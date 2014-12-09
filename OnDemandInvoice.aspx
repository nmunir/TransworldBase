<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Dim gsInvoiceGUID As String
    Dim gdtOrderData As DataTable
    Dim gdtInvoiceItems As DataTable
    Dim gdtInvoice As DataTable
    Dim gdblNetTotal As Double
    Dim gdblVATRate As Double
    Dim gdblVATTotal As Double
    Dim gdblGrossTotal As Double
    Dim gbDraft As Boolean = False
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Try
                Try
                    gsInvoiceGUID = Request.QueryString("Key")
                    If gsInvoiceGUID.Trim = String.Empty Then
                        WebMsgBox.Show("Bad parameter (1)")
                        Exit Sub
                    End If
                Catch ex As Exception
                    WebMsgBox.Show("Bad parameter (2)" & ex.Message)
                    Exit Sub
                End Try
                Try
                    Dim sDraftFlag As String = Request.QueryString("Flag")
                    If sDraftFlag <> String.Empty Then
                        gbDraft = True
                    End If
                Catch ex As Exception
                    WebMsgBox.Show("Bad parameter (3)" & ex.Message)
                    Exit Sub
                End Try
                Call GetOrderData()
                If gdtOrderData.Rows.Count <> 1 Then
                    WebMsgBox.Show("Could not find referenced invoice.")
                    Exit Sub
                End If
                lblInvoiceNumber.Text = gdtOrderData.Rows(0).Item("id")
                Call BuildInvoiceAddress()
                Call BuildDeliveryAddress()

                gdblVATRate = gdtOrderData.Rows(0).Item("VATRate")
                lblLegendVAT.Text = "VAT @ " & gdblVATRate & "%"

                gdtInvoice = BuildInvoiceItemsTable()
                gdblNetTotal = 0

                If gbDraft Then
                    lblDraft01.Visible = True
                    lblDraft02.Visible = True
                    lblInvoiceDate.Text = "Order placed " & Format(DateTime.Now, "dd MMM yyyy h:mm") & " by " & Session("UserName")
                    For Each drInvoiceItem As DataRow In gdtInvoiceItems.Rows
                        Dim dr As DataRow
                        dr = gdtInvoice.NewRow
                        dr("Product") = GetProductDetails(drInvoiceItem("ProductKey"))
                        dr("QuantityOrdered") = drInvoiceItem("QuantityOrdered")
                        dr("PricePerItem") = GetPricePerItem(drInvoiceItem("TariffId"), drInvoiceItem("QuantityOrdered"))
                        dr("TotalPrice") = CInt(drInvoiceItem("QuantityOrdered")) * CDbl(dr("PricePerItem"))
                        gdtInvoice.Rows.Add(dr)
                        gdblNetTotal += dr("TotalPrice")
                    Next
                Else
                    lblInvoiceDate.Text = "Order placed " & GetOrderDate() & " by " & GetOrdererName()
                    For Each drInvoiceItem As DataRow In gdtInvoiceItems.Rows
                        Dim dr As DataRow
                        dr = gdtInvoice.NewRow
                        dr("Product") = GetProductDetails(drInvoiceItem("ProductKey"))
                        dr("QuantityOrdered") = drInvoiceItem("QuantityOrdered")
                        dr("PricePerItem") = drInvoiceItem("PricePerItem")
                        dr("TotalPrice") = drInvoiceItem("TotalPrice")
                        gdtInvoice.Rows.Add(dr)
                        gdblNetTotal += dr("TotalPrice")
                    Next
                End If
                gvInvoiceItems.DataSource = gdtInvoice
                gvInvoiceItems.DataBind()
            Catch ex As Exception
                WebMsgBox.Show("Error - " & ex.Message)
            End Try
            
            lblPurchaseOrderNo.Text = gdtOrderData.Rows(0).Item("POCode") & String.Empty
            
            gdblVATTotal = (gdblNetTotal * gdblVATRate) / 100
            gdblGrossTotal = gdblNetTotal + gdblVATTotal
            lblNetTotal.Text = Format(gdblNetTotal, "£##,##0.00")
            lblVATAmount.Text = Format(gdblVATTotal, "£##,##0.00")
            lblGrossTotal.Text = Format(gdblGrossTotal, "£##,##0.00")
        End If
    End Sub
    
    Protected Function GetOrderDate() As String
        GetOrderDate = Format(DateTime.Parse(ExecuteQueryToDataTable("SELECT CreatedOn FROM Consignment WHERE [key] = " & gdtOrderData.Rows(0).Item("ConsignmentKey")).Rows(0).Item(0) & String.Empty), "dd MMM yyyy h:mm")
    End Function

    Protected Function GetOrdererName() As String
        GetOrdererName = ExecuteQueryToDataTable("SELECT FirstName + ' ' + LastName FROM UserProfile up INNER JOIN Consignment c ON c.UserKey = up.[key] WHERE c.[key] = " & gdtOrderData.Rows(0).Item("ConsignmentKey")).Rows(0).Item(0) & String.Empty
    End Function

    Protected Sub GetOrderData()
        Dim sSQL As String
        sSQL = "SELECT * FROM OnDemandTransaction WHERE SessionGUID = '" & gsInvoiceGUID & "'"
        gdtOrderData = ExecuteQueryToDataTable(sSQL)
        sSQL = "SELECT * FROM OnDemandTransactionStatus WHERE SessionGUID = '" & gsInvoiceGUID & "'"
        gdtInvoiceItems = ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function GetPricePerItem(ByVal nTariffId As Integer, ByVal nQuantityOrdered As Integer) As Double
        Dim sSQL As String
        sSQL = "SELECT * FROM OnDemandTariff WHERE TariffId = " & nTariffId & " ORDER BY Quantity"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim dblLastPrice As Double = -1
        For Each dr As DataRow In oDataTable.Rows
            dblLastPrice = dr("Price")
            If nQuantityOrdered <= dr("Quantity") Then
                Exit For
            End If
        Next
        GetPricePerItem = dblLastPrice
    End Function
    
    Protected Sub BuildInvoiceAddress()
        Dim sbAddress As New StringBuilder
        Dim dr As DataRow = gdtOrderData.Rows(0)
        If dr("InvoiceAttnOf") <> String.Empty Then
            sbAddress.Append(dr("InvoiceAttnOf"))
            sbAddress.Append("<br />")
        End If

        sbAddress.Append(dr("InvoiceRecipient"))
        sbAddress.Append("<br />")

        sbAddress.Append(dr("InvoiceAddress1"))
        sbAddress.Append("<br />")

        If dr("InvoiceAddress1") <> String.Empty Then
            sbAddress.Append(dr("InvoiceAddress2"))
            sbAddress.Append("<br />")
        End If
        If dr("InvoiceAddress3") <> String.Empty Then
            sbAddress.Append(dr("InvoiceAddress3"))
            sbAddress.Append("<br />")
        End If

        sbAddress.Append(dr("InvoiceTown"))
        sbAddress.Append("<br />")

        If dr("InvoicePostcode") <> String.Empty Then
            sbAddress.Append(dr("InvoicePostcode"))
            sbAddress.Append("<br />")
        End If
        lblInvoiceAddress.Text = sbAddress.ToString
    End Sub
    
    Protected Sub BuildDeliveryAddress()
        Dim sbAddress As New StringBuilder
        Dim dr As DataRow = gdtOrderData.Rows(0)
        If dr("DeliveryAttnOf") <> String.Empty Then
            sbAddress.Append(dr("DeliveryAttnOf"))
            sbAddress.Append("<br />")
        End If

        sbAddress.Append(dr("DeliveryRecipient"))
        sbAddress.Append("<br />")

        sbAddress.Append(dr("DeliveryAddress1"))
        sbAddress.Append("<br />")

        If dr("DeliveryAddress2") <> String.Empty Then
            sbAddress.Append(dr("DeliveryAddress2"))
            sbAddress.Append("<br />")
        End If
        If dr("DeliveryAddress3") <> String.Empty Then
            sbAddress.Append(dr("DeliveryAddress3"))
            sbAddress.Append("<br />")
        End If

        sbAddress.Append(dr("DeliveryTown"))
        sbAddress.Append("<br />")

        If dr("DeliveryPostcode") <> String.Empty Then
            sbAddress.Append(dr("DeliveryPostcode"))
            sbAddress.Append("<br />")
        End If
        lblDeliveryAddress.Text = sbAddress.ToString
    End Sub
    
    Protected Function BuildInvoiceItemsTable() As DataTable
        Dim dtInvoiceItems As New DataTable
        dtInvoiceItems.Columns.Add(New DataColumn("Product", GetType(String)))
        dtInvoiceItems.Columns.Add(New DataColumn("QuantityOrdered", GetType(Integer)))
        dtInvoiceItems.Columns.Add(New DataColumn("PricePerItem", GetType(Double)))
        dtInvoiceItems.Columns.Add(New DataColumn("TotalPrice", GetType(Double)))
        BuildInvoiceItemsTable = dtInvoiceItems
    End Function

    Protected Function GetProductDetails(ByVal nProductKey As Integer) As String
        Dim sProductDetails As String = String.Empty
        Dim sSQL As String = "SELECT ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & nProductKey
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim dr As DataRow = oDataTable.Rows(0)
        sProductDetails = dr("ProductCode")
        sProductDetails += " - "
        Dim sProductDate As String = dr("ProductDate") & String.Empty
        If sProductDate <> String.Empty Then
            sProductDetails += sProductDate & " - "
        End If
        sProductDetails += dr("ProductDescription") & String.Empty
        GetProductDetails = sProductDetails
    End Function
    
    Protected Function ExecuteQueryToListItemCollection(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If Not IsDBNull(oDataReader(sTextFieldName)) Then
                        sTextField = oDataReader(sTextFieldName)
                    Else
                        sTextField = String.Empty
                    End If
                    If Not IsDBNull(oDataReader(sValueFieldName)) Then
                        sValueField = oDataReader(sValueFieldName)
                    Else
                        sValueField = String.Empty
                    End If
                    oListItemCollection.Add(New ListItem(sTextField, sValueField))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQuery: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
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

    Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
        ExecuteNonQuery = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            oCmd.ExecuteNonQuery()
            ExecuteNonQuery = True
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Invoice</title>
    <style type="text/css">
    body {
            background:#ffffff;
            color:#666666;
            font-family:verdana
    }
        p { font-size:100%;color:#000000; }
        h1 { font-size:150%; color:#000099; }

        .style1
        {
            width: 369px;
        }
        .style2
        {
            width: 161px;
        }

    </style>
</head>
<body>
    <form id="frmOnDemandPriceList" runat="server">
        <table style="width: 100%">
            <tr>
                <td style="width: 60%; white-space: nowrap" align="right">
                    <img src="images/logos/transworld.jpg" alt="" />&nbsp;</td>
                <td style="width: 39%" valign="top"><div style="font-family:Verdana; font-size:xx-small">
                    <b>Complete Marketing Support</b><br />
                    Storage and Distribution<br />
                    Fulfilment, UK and International Mail<br />
                    UK and Worldwide Courier<br />
                    Call Centre Services<br />
                    Design and Print</div></td>
                <td style="width: 40%" valign="top">
                    <asp:Button ID="btnCloseWindow" OnClientClick="window.close()" runat="server" Text="close window" /></td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 90%; white-space: nowrap">
                    &nbsp;<asp:Label ID="lblDraft01" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="D R A F T&nbsp;&nbsp;" Visible="False"></asp:Label><asp:Label ID="Label6" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="I N V O I C E"></asp:Label>&nbsp;
                    <asp:Label ID="lblInvoiceNumber" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Visible="False"/>
                </td>
                <td style="width: 10%" align="right">
                    &nbsp;</td>
            </tr>
            <tr>
                <td style="white-space: nowrap">
                    <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Invoice To:"/>
                    </td>
                <td align="right"/>
            </tr>
            <tr>
                <td style="white-space: nowrap; height: 53px;"><asp:Label ID="lblInvoiceAddress" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"/><br />
                    <br />
                    <asp:Label ID="Label5" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Deliver To:"/><br />
                    <asp:Label ID="lblDeliveryAddress" runat="server" Font-Bold="True" Font-Size="X-Small"/><br />
                    <br />
                    <asp:Label ID="lblInvoiceDate" runat="server" Font-Bold="True" Font-Size="X-Small"/><br />
                    </td>
                <td align="right" style="height: 53px"/>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="lblPriceList" runat="server" Font-Size="XX-Small"></asp:Label>
                    <br />
                    <asp:GridView ID="gvInvoiceItems" runat="server" Font-Size="XX-Small" AutoGenerateColumns="False" Width="95%">
                        <Columns>
                            <asp:TemplateField HeaderText="Product">
                                <ItemTemplate>
                                    <asp:Label ID="lblProduct" runat="server" Text='<%# Container.DataItem("Product") %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Quantity Ordered">
                                <ItemTemplate>
                                    <asp:Label ID="lblQuantityOrdered" runat="server" Text='<%# Container.DataItem("QuantityOrdered") %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Price per Item (£)">
                                <ItemTemplate>
                                    <asp:Label ID="lblPricePerItem" runat="server" Text='<%# Format(Container.DataItem("PricePerItem"),"##,##0.00") %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Total Price (£)">
                                <ItemTemplate>
                                    <asp:Label ID="lblTotalPrice" runat="server" Text='<%# Format(Container.DataItem("TotalPrice"),"##,##0.00") %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="height: 72px">
                    <br />
                    <br />
                    <br />
                    <table class="style1">
                        <tr>
                            <td class="style2">
                    <asp:Label ID="Label7" runat="server" Font-Bold="True" Font-Names="Verdana" 
                                    Font-Size="X-Small" Text="Purchase Order #:"/>
                            </td>
                            <td>
                    <asp:Label ID="lblPurchaseOrderNo" runat="server" Font-Bold="True" Font-Size="X-Small"/>
                            </td>
                        </tr>
                        <tr>
                            <td class="style2">
                                &nbsp;</td>
                            <td>
                                &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style2">
                    <asp:Label ID="lblLegendNetAmount" runat="server" Font-Bold="True" Font-Size="X-Small" Text="Net Amount: "/>
                            </td>
                            <td>
                    <asp:Label ID="lblNetTotal" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="X-Small" ForeColor="Black"/>
                            </td>
                        </tr>
                        <tr>
                            <td class="style2">
                                <asp:Label ID="lblLegendVAT" runat="server" Font-Bold="True" Font-Size="X-Small" Text="VAT: "/>
                            </td>
                            <td>
                    <asp:Label ID="lblVATAmount" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="X-Small" ForeColor="Black"/>
                            </td>
                        </tr>
                        <tr>
                            <td class="style2">
                    <asp:Label ID="lblLegendGrossAmount" runat="server" Font-Bold="True" Font-Size="X-Small" 
                                    Text="Total: "/>
                            </td>
                            <td>
                    <asp:Label ID="lblGrossTotal" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="X-Small" ForeColor="Black"/>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <div style="font-family:Verdana; font-size:xx-small">
                    Transworld Marketing Logistics Ltd.
                    The Mercury Centre Central Way&nbsp; Fetham&nbsp; Middlesex TW14 0RN<br />
                    Tel: 020 851 1111 Fax: 020 8890 9090 Email: sales@transworld.eu.com&nbsp; www.transworld.eu.com<br />
                    <br />
                    Registered in England and Wales&nbsp; Registration number 2314301&nbsp; VAT number
                    530 2969 51<br />
                    Part of Badr Logistics Ltd.</div><br />
                    <br />
                    <asp:Label ID="lblDraft02" runat="server" Font-Bold="True" Font-Size="XX-Small" Visible="False">PLEASE NOTE: this is a DRAFT invoice. A link to your final invoice, which will be identical to this draft if you complete your order without further modification, will be emailed to you on completion of your order.</asp:Label></td>
                <td style="height: 72px" />
            </tr>
        </table>
    </form>
</body>
</html>
