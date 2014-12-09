<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Drawing.Image" %>
<%@ import Namespace="System.Drawing.Color" %>
<script runat="server">
    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")

    Private Shared gdvExportData As DataView
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            GetYears()
            ShowYearSelection()
        End If
        If Not IsNumeric(Session("CustomerKey")) Then
            'Server.Transfer("../session_expired.aspx")
        End If
    End Sub
    
    Protected Sub ShowYearSelection()
        pnlYearSelection.Visible = True
        pnlInvoiceList.Visible = False
        If Repeater1.Visible = False Then
            lblSelectLegend.Text = "No data available"
        End If
    End Sub
    
    Protected Sub ShowInvoiceList()
        pnlYearSelection.Visible = False
        pnlInvoiceList.Visible = True
    End Sub
    
    Protected Sub btn_ShowMonths_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psYear = CStr(e.CommandArgument)
        lblYearHeader.Text = psYear
        Repeater2.Visible = True
        GetMonths()
    End Sub
    
    Protected Sub btn_ShowInvoices_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psMonth = CStr(e.CommandArgument)
        lblMonthHeader.Text = psMonth
        BindInvoiceGrid("ProductDescription")
        ShowInvoiceList()

        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
    End Sub
    
    Protected Sub repeater1_item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = System.Drawing.Color.Blue
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = System.Drawing.Color.Red
    End Sub
    
    Protected Sub btn_ReSelectYear_click(ByVal s As Object, ByVal e As EventArgs)
        Dim item As RepeaterItem
        For Each item In Repeater1.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = System.Drawing.Color.Blue
        Next
        Repeater2.Visible = False
        ShowYearSelection()
    End Sub
    
    Protected Sub btn_DownloadCSVFile_Click(ByVal sender As Object, ByVal e As EventArgs)
        ExportCSVData()
    End Sub
    
    Protected Sub BindInvoiceGrid(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Hyster_GetInvoices", oConn)
        lblInvoiceMessage.Text = ""
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Year", SqlDbType.NChar, 4))
        oAdapter.SelectCommand.Parameters("@Year").Value = psYear
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Month", SqlDbType.NVarChar, 15))
        oAdapter.SelectCommand.Parameters("@Month").Value = psMonth
        lblInvoiceMessage.Text = ""
        oAdapter.Fill(oDataSet, "Invoices")
        'Dim Source As DataView = oDataSet.Tables("Invoices").DefaultView
        pdvInvoiceDataView = oDataSet.Tables("Invoices").DefaultView
        'Source.Sort = SortField
        pdvInvoiceDataView.Sort = SortField
        'If Source.Count > 0 Then
        If pdvInvoiceDataView.Count > 0 Then
            'dgrdInvoices.DataSource = Source
            dgrdInvoices.DataSource = pdvInvoiceDataView
            dgrdInvoices.DataBind()
            dgrdInvoices.Visible = True
        Else
            lblInvoiceMessage.Text = "No invoices found"
            dgrdInvoices.Visible = False
        End If
        oConn.Close()
    End Sub
    
    Protected Sub SortInvoiceColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        BindInvoiceGrid(E.SortExpression)
    End Sub
    
    Protected Sub GetYears()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Hyster_GetInvoiceYears", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "Years")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                Repeater1.Visible = True
                Repeater1.DataSource = oDataSet
                Repeater1.DataBind()
            Else
                Repeater1.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetMonths()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Hyster_GetInvoiceMonths", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Year", SqlDbType.NChar, 4))
        oAdapter.SelectCommand.Parameters("@Year").Value = psYear
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "Months")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                Repeater2.Visible = True
                Repeater2.DataSource = oDataSet
                Repeater2.DataBind()
            Else
                Repeater2.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ExportCSVData()                 ' called from button_click; data to be exported must be in a gdv
        Response.Clear()                        ' ie Dim response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "SprintInvoiceData.csv")
        'Response.ContentType = "application/vnd.ms-excel"
        Response.ContentType = "text/csv"
    
        Dim sExportContent As String = String.Empty
        If (Not gdvExportData Is Nothing) AndAlso gdvExportData.Table.Rows.Count > 0 Then
            sExportContent = ConvertDataViewToCSVString(gdvExportData)
        End If
        If sExportContent.Length <= 0 Then
            sExportContent = "Data not found"
        End If
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sExportContent)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.Flush()  ' may not be necessary but won't hurt
        Response.End()
    End Sub
    
    Public Function ConvertDataViewToCSVString(ByVal oDataView As DataView) As String
        Dim sSingleQuote As String = """"
        Dim sDoubleQuote As String = sSingleQuote & sSingleQuote
        Dim ResultBuilder As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRowView

        For Each oDataColumn In oDataView.Table.Columns         ' column headings in line 1
            ResultBuilder.Append(oDataColumn.ColumnName)
            ResultBuilder.Append(",")
        Next
        If ResultBuilder.Length > 1 Then
            ResultBuilder.Length = ResultBuilder.Length - 1
        End If
        ResultBuilder.Append(Environment.NewLine)
    
        For Each oDataRow In oDataView
            For Each oDataColumn In oDataView.Table.Columns
                Dim x As String = String.Empty
                If Not IsDBNull(oDataRow(oDataColumn.ColumnName)) Then
                    x = oDataRow(oDataColumn.ColumnName)
                    x = x.Replace(sSingleQuote, sDoubleQuote)
                    If x.Contains(sSingleQuote) Then
                        Dim y = x
                    End If
                End If
                x = sSingleQuote & x & sSingleQuote
                x += ","
                ResultBuilder.Append(x)
            Next oDataColumn
            ResultBuilder.Length = ResultBuilder.Length - 1
            ResultBuilder.Append(Environment.NewLine)
        Next oDataRow

        If Not ResultBuilder Is Nothing Then
            Return ResultBuilder.ToString()
        Else
            Return String.Empty
        End If
    End Function

    Property pdvInvoiceDataView() As DataView
        Get
            Return gdvExportData
        End Get
    
        Set(ByVal Value As DataView)
            gdvExportData = Value
        End Set
    End Property
    
    Property psYear() As String
        Get
            Dim o As Object = ViewState("HIR_Year")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HIR_Year") = Value
        End Set
    End Property
    
    Property psMonth() As String
        Get
            Dim o As Object = ViewState("HIR_Month")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("HIR_Month") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Hyster Invoice Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form id="HysterInvoiceReport" runat="server">
          <asp:Table id="TableHeader" runat="server" width="100%">
              <asp:TableRow>
                  <asp:TableCell VerticalAlign="Bottom" width="0%"></asp:TableCell>
                  <asp:TableCell Wrap="False" width="50%">
                      <asp:Label ID="Label1" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial">Hyster
                      Invoice Report</asp:Label><br /><br />
                  </asp:TableCell>
                  <asp:TableCell Wrap="False" HorizontalAlign="Right" width="50%"></asp:TableCell>
              </asp:TableRow>
          </asp:Table>
        <asp:Panel id="pnlYearSelection" runat="server" CellSpacing="0" visible="True">
            <asp:Table id="tblYearSelection" runat="server" font-names="Verdana" Font-Size="Small" Width="100%">
                <asp:TableRow>
                    <asp:TableCell Width="5%"></asp:TableCell>
                    <asp:TableCell Width="50%"></asp:TableCell>
                    <asp:TableCell Width="45%"></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell Wrap="False" ColumnSpan="2">
                        <br />
                        <asp:Label ID="lblSelectLegend" runat="server" font-size="X-Small">Select invoice year, then month</asp:Label>
                        <br />
                        <br />
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                        <asp:Repeater runat="server" ID="Repeater1" OnItemCommand="repeater1_item_click">
                            <ItemTemplate>
                                <asp:Image runat="server" ImageUrl="../images/icon_arrow.gif"></asp:Image>
                                <asp:LinkButton runat="server" ForeColor="Blue" OnCommand="btn_ShowMonths_click" CommandArgument='<%# Container.DataItem("InvoicingYear")%>' Text='<%# Container.DataItem("InvoicingYear")%>'></asp:LinkButton>
                                <br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                        <asp:Repeater runat="server" Visible="False" ID="Repeater2">
                            <ItemTemplate>
                                <asp:Image runat="server" ImageUrl="../images/icon_arrow.gif"></asp:Image>
                                <asp:LinkButton runat="server" ForeColor="Blue" OnCommand="btn_ShowInvoices_click" CommandArgument='<%# Container.DataItem("InvoicingMonthName")%>' Text='<%# Container.DataItem("InvoicingMonthName")%>'></asp:LinkButton>
                                <br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Panel id="pnlInvoiceList" runat="server" CellSpacing="0" visible="True">
            <asp:table id="Table3" runat="Server" width="90%">
                <asp:TableRow >
                    <asp:TableCell Wrap="False" ColumnSpan="2">
                        <asp:Label runat="server" id="lblYearHeader" font-names="Verdana" forecolor="Blue" font-bold="True" font-size="Small"></asp:Label>
                        &nbsp;<asp:Label runat="server" id="lblMonthHeader" font-names="Verdana" forecolor="Blue" font-bold="True" font-size="Small"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow >
                <asp:TableRow >
                    <asp:TableCell VerticalAlign="Bottom" width="5%">
                        <asp:Image runat="server" ImageUrl="../images/icon_back.gif"></asp:Image>
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" VerticalAlign="Top">
                        <asp:LinkButton runat="server" ForeColor="Blue" Font-Size="X-Small" Font-Names="Verdana" OnClick="btn_ReSelectYear_click">re-select&nbsp;period</asp:LinkButton>
                        &nbsp;&nbsp;<asp:Button  Text="download to excel" onclick="btn_DownloadCSVFile_Click" runat="server" Tooltip="download data as a CSV file"></asp:Button>
                    </asp:TableCell>
                </asp:TableRow >
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                </asp:TableRow>
            </asp:table>
            <asp:Label id="lblInvoiceMessage" runat="server" font-names="Verdana" font-size="X-Small" forecolor="#00C000"></asp:Label>
            <asp:DataGrid id="dgrdInvoices" runat="server" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" OnSortCommand="SortInvoiceColumns" AllowSorting="True" ShowFooter="True" GridLines="None" Font-Names="Verdana">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" wrap="False" bordercolor="Gray"></HeaderStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle font-names="Verdana" backcolor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="DealerOrderDate" SortExpression="DealerOrderDate" HeaderText="Dealer Order Date">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="SprintJobNo" SortExpression="SprintJobNo" HeaderText="Sprint Job No">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ConsignmentNo" SortExpression="ConsignmentNo" HeaderText="Consignment No">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="DealerCompanyName" SortExpression="DealerCompanyName" HeaderText="Dealer Company Name">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="DealerCountry" SortExpression="DealerCountry" HeaderText="Dealer Country">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BranchDepartmentCode" SortExpression="BranchDepartmentCode" HeaderText="Branch Department Code">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="NOMsDealerCode" SortExpression="NOMsDealerCode" HeaderText="NOMs Dealer Code">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PONo" SortExpression="PONo" HeaderText="PO No">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Product Code">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription" HeaderText="Product Description">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PartType" SortExpression="PartType" HeaderText="Part Type">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Quantity" SortExpression="Quantity" HeaderText="Quantity">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="UnitPrice" SortExpression="UnitPrice" HeaderText="Unit Price" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PriceCurrency" SortExpression="PriceCurrency" HeaderText="Price Currency">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ActualShipmentDate" SortExpression="ActualShipmentDate" HeaderText="Actual Shipment Date">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CarriageCost" SortExpression="CarriageCost" HeaderText="Carriage Cost" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CarriageCurrency" SortExpression="CarriageCurrency" HeaderText="Carriage Currency">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine1" SortExpression="AddressLine1" HeaderText="Address Line 1">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine2" SortExpression="AddressLine2" HeaderText="Address Line 2">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine3" SortExpression="AddressLine3" HeaderText="Address Line 3">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ManualFlag" SortExpression="ManualFlag" HeaderText="Manual Flag">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Brand" SortExpression="Brand" HeaderText="Brand">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="EmailAddress" SortExpression="EmailAddress" HeaderText="Email Address">
                        <HeaderStyle wrap="False" horizontalalign="Left" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <asp:Table id="Table4" runat="Server" Font-Size="X-Small" Width="100%" Font-Name="Verdana">
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right"></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
          <asp:Label ID="lblReportGeneratedDateTime" Visible="True" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green"></asp:Label>
        </asp:Panel>
        <asp:Label id="lblError" runat="server" font-names="Verdana" font-size="X-Small" forecolor="red"></asp:Label>
    </form>
</body>
</html>
