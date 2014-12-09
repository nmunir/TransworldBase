<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>

<script runat="server">

    ' BlackRock Invoice Report

    ' TO DO
    ' Get list of fields to be displayed from GG
    ' Clone Hyster batch job(s) as necessary

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private Shared gdvExportData As DataView
    Private Const C_ERROR_NO_RESULT As String = "Data not found"
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            GetYears()
            ShowYearSelection()
        End If
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlYearSelection.Visible = False
        pnlInvoiceList.Visible = False
    End Sub
    
    Protected Sub ShowYearSelection()
        Call HideAllPanels()
        pnlYearSelection.Visible = True
    End Sub
    
    Protected Sub ShowInvoiceList()
        Call HideAllPanels()
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
        Dim lnkbtnInvoiceYear As LinkButton
        For Each item In s.Items
            lnkbtnInvoiceYear = CType(item.Controls(3), LinkButton)
            lnkbtnInvoiceYear.ForeColor = System.Drawing.Color.Blue
        Next
        lnkbtnInvoiceYear = CType(e.CommandSource, LinkButton)
        lnkbtnInvoiceYear.ForeColor = System.Drawing.Color.Red
    End Sub
    
    Protected Sub btn_ReSelectYear_click(ByVal s As Object, ByVal e As EventArgs)
        Dim ri As RepeaterItem
        For Each ri In Repeater1.Items
            Dim lb As LinkButton = CType(ri.Controls(3), LinkButton)
            lb.ForeColor = System.Drawing.Color.Blue
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
        Dim oAdapter As New SqlDataAdapter("spASPNET_BlackRock_GetInvoices", oConn)
        lblInvoiceMessage.Text = ""
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Year", SqlDbType.NChar, 4))
        oAdapter.SelectCommand.Parameters("@Year").Value = psYear
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Month", SqlDbType.NVarChar, 15))
        oAdapter.SelectCommand.Parameters("@Month").Value = psMonth
        lblInvoiceMessage.Text = ""
        oAdapter.Fill(oDataSet, "Invoices")
        pdvInvoiceDataView = oDataSet.Tables("Invoices").DefaultView
        pdvInvoiceDataView.Sort = SortField
        If pdvInvoiceDataView.Count > 0 Then
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
        Dim oAdapter As New SqlDataAdapter("spASPNET_BlackRock_GetInvoiceYears", oConn)
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
        Dim oAdapter As New SqlDataAdapter("spASPNET_BlackRock_GetInvoiceMonths", oConn)
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
    
    Private Sub ExportCSVData()
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "SprintInvoiceData.csv")   ' Add the header that specifies the default filename for the Download/SaveAs dialog
        Response.ContentType = "application/vnd.ms-excel"                                             ' Specify that the response is a stream that cannot be read by the client and must be downloaded
    
        Dim sExportContent As String = String.Empty
        If (Not gdvExportData Is Nothing) AndAlso gdvExportData.Table.Rows.Count > 0 Then
            sExportContent = ConvertDataViewToCSVString(gdvExportData)
        End If
        If sExportContent.Length <= 0 Then
            sExportContent = C_ERROR_NO_RESULT
        End If
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sExportContent)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.Flush()
        Response.End()          ' Stop execution of the current page
    End Sub
    
    Public Function ConvertDataViewToCSVString(ByVal oDataView As DataView) As String
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
                ResultBuilder.Append(oDataRow(Replace(oDataColumn.ColumnName, ",", " ")))  ' replace any commas with spaces
                ResultBuilder.Append(",")
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

    Property psYear() As String
        Get
            Dim o As Object = ViewState("Year")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("Year") = Value
        End Set
    End Property
     
    Property psMonth() As String
        Get
            Dim o As Object = ViewState("Month")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("Month") = Value
        End Set
    End Property
    
    Property pdvInvoiceDataView() As DataView
        Get
            Return gdvExportData
        End Get
    
        Set(ByVal Value As DataView)
            gdvExportData = Value
        End Set
    End Property
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>BlackRock Invoice Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form runat="server">
        <asp:Table ID="TableHeader" runat="server" Width="100%">
            <asp:TableRow>
                <asp:TableCell VerticalAlign="Bottom" Width="0%"></asp:TableCell>
                <asp:TableCell Wrap="False" Width="50%">
                    <asp:Label ID="Label1" runat="server" ForeColor="silver" Font-Size="Small" Font-Bold="True"
                        Font-Names="Arial" Text="BlackRock Invoice Report" /><br />
                    <br />
                </asp:TableCell>
                <asp:TableCell Wrap="False" HorizontalAlign="Right" Width="50%"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        <asp:Panel ID="pnlYearSelection" runat="server" CellSpacing="0" Visible="True">
            <asp:Table ID="tblYearSelection" runat="server" Font-Names="Verdana" Font-Size="Small"
                Width="100%">
                <asp:TableRow>
                    <asp:TableCell Width="5%"></asp:TableCell>
                    <asp:TableCell Width="50%"></asp:TableCell>
                    <asp:TableCell Width="45%"></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell Wrap="False" ColumnSpan="2">
                        <br />
                        <asp:Label runat="server" font-size="X-Small">Select invoice year, then month</asp:Label>
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
                                <asp:LinkButton ID="lnkbtnInvoiceYear" runat="server" ForeColor="Blue" OnCommand="btn_ShowMonths_click"
                                    CommandArgument='<%# Container.DataItem("InvoicingYear")%>' Text='<%# Container.DataItem("InvoicingYear")%>'></asp:LinkButton>
                                <br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                        <asp:Repeater runat="server" Visible="False" ID="Repeater2">
                            <ItemTemplate>
                                <asp:Image runat="server" ImageUrl="../images/icon_arrow.gif"></asp:Image>
                                <asp:LinkButton ID="lnkbtnInvoiceMonth" runat="server" ForeColor="Blue" OnCommand="btn_ShowInvoices_click"
                                    CommandArgument='<%# Container.DataItem("InvoicingMonthName")%>' Text='<%# Container.DataItem("InvoicingMonthName")%>'></asp:LinkButton>
                                <br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Panel ID="pnlInvoiceList" runat="server" CellSpacing="0" Visible="True">
            <asp:Table ID="Table3" runat="Server" Width="90%">
                <asp:TableRow>
                    <asp:TableCell Wrap="False" ColumnSpan="2">
                        <asp:Label runat="server" ID="lblYearHeader" Font-Names="Verdana" ForeColor="Blue"
                            Font-Bold="True" Font-Size="Small"></asp:Label>
                        &nbsp;<asp:Label runat="server" ID="lblMonthHeader" Font-Names="Verdana" ForeColor="Blue"
                            Font-Bold="True" Font-Size="Small"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell VerticalAlign="Bottom" Width="5%">
                        <asp:Image runat="server" ImageUrl="../images/icon_back.gif"></asp:Image>
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" VerticalAlign="Top">
                        <asp:LinkButton runat="server" ForeColor="Blue" Font-Size="X-Small" Font-Names="Verdana" OnClick="btn_ReSelectYear_click">re-select&nbsp;period</asp:LinkButton>
                        &nbsp;&nbsp;<asp:Button  Text="download to Excel" onclick="btn_DownloadCSVFile_Click" runat="server" Tooltip="download data as a CSV file"></asp:Button>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Label ID="lblInvoiceMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"
                ForeColor="#00C000"></asp:Label>
            <asp:DataGrid ID="dgrdInvoices" runat="server" Font-Size="XX-Small" Width="100%"
                AutoGenerateColumns="False" OnSortCommand="SortInvoiceColumns" AllowSorting="True"
                ShowFooter="True" GridLines="None" Font-Names="Verdana">
                <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" BorderColor="Gray">
                </HeaderStyle>
                <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                <ItemStyle Font-Names="Verdana" BackColor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="DealerOrderDate" SortExpression="DealerOrderDate" HeaderText="Dealer Order Date">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="SprintJobNo" SortExpression="SprintJobNo" HeaderText="Sprint Job No">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ConsignmentNo" SortExpression="ConsignmentNo" HeaderText="Consignment No">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="DealerCompanyName" SortExpression="DealerCompanyName"
                        HeaderText="Dealer Company Name">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="DealerCountry" SortExpression="DealerCountry" HeaderText="Dealer Country">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BranchDepartmentCode" SortExpression="BranchDepartmentCode"
                        HeaderText="Branch Department Code">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="NOMsDealerCode" SortExpression="NOMsDealerCode" HeaderText="NOMs Dealer Code">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PONo" SortExpression="PONo" HeaderText="PO No">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Product Code">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription"
                        HeaderText="Product Description">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PartType" SortExpression="PartType" HeaderText="Part Type">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Quantity" SortExpression="Quantity" HeaderText="Quantity">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="UnitPrice" SortExpression="UnitPrice" HeaderText="Unit Price"
                        DataFormatString="{0:#,##0.00}">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PriceCurrency" SortExpression="PriceCurrency" HeaderText="Price Currency">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ActualShipmentDate" SortExpression="ActualShipmentDate"
                        HeaderText="Actual Shipment Date">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CarriageCost" SortExpression="CarriageCost" HeaderText="Carriage Cost"
                        DataFormatString="{0:#,##0.00}">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CarriageCurrency" SortExpression="CarriageCurrency" HeaderText="Carriage Currency">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine1" SortExpression="AddressLine1" HeaderText="Address Line 1">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine2" SortExpression="AddressLine2" HeaderText="Address Line 2">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AddressLine3" SortExpression="AddressLine3" HeaderText="Address Line 3">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ManualFlag" SortExpression="ManualFlag" HeaderText="Manual Flag">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Brand" SortExpression="Brand" HeaderText="Brand">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="EmailAddress" SortExpression="EmailAddress" HeaderText="Email Address">
                        <HeaderStyle Wrap="False" HorizontalAlign="Left" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <asp:Table ID="Table4" runat="Server" Font-Size="X-Small" Width="100%" Font-Name="Verdana">
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right"></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Label ID="lblReportGeneratedDateTime" Visible="True" runat="server" Text=""
                Font-Size="XX-Small" Font-Names="Verdana, Sans-Serif" ForeColor="Green"></asp:Label>
        </asp:Panel>
        <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="red"></asp:Label>
    </form>
</body>
</html>
