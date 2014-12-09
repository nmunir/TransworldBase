<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '   Product Value Report

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
        If Not IsPostBack Then
            pbIsProductOwner = CBool(Session("UserType").ToString.ToLower.Contains("owner"))
            Call GetSiteFeatures()
            trProductGroups.Visible = pbProductOwners
            ' pbProductOwners = site-wide Product Owners attribute; pbIsProductOwner = this user
            If pbIsProductOwner Then
                If pbProductOwners Then
                    ddlProductGroup.Visible = True
                    PopulateProductGroups(Session("UserKey"))
                    btnShowProductGroups.Visible = False
                Else
                    WebMsgBox.Show("Cannot show report as Product Owners attribute is not enabled for this web site")
                    Exit Sub
                End If
            Else
                If pbProductOwners Then
                    btnShowProductGroups.Visible = True
                Else
                    btnShowProductGroups.Visible = False
                End If
                pnSelectedProductGroup = 0
            End If
            Call ShowProductSelection()
        End If
    End Sub
    
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettings"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pbProductOwners = dr("ProductOwners")
    End Sub
    
    Protected Sub PopulateProductGroups(ByVal nProductOwner As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetGroupsForOwner", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramProductOwner As SqlParameter = New SqlParameter("@ProductOwner", SqlDbType.Int)
        paramProductOwner.Value = nProductOwner
        oCmd.Parameters.Add(paramProductOwner)
       
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
       
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                ddlProductGroup.Items.Add(New ListItem("- select product group -", -1))
                If Not pbIsProductOwner Then
                    ddlProductGroup.Items.Add(New ListItem("- all products -", 0))
                End If
                While oSqlDataReader.Read()
                    ddlProductGroup.Items.Add(New ListItem(oSqlDataReader("ProductGroupName"), oSqlDataReader("ProductGroupKey")))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("PopulateProductgGroupsDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        If ddlProductGroup.Items.Count <= 2 Then
            lblProductGroup.Text = "Product group: " & ddlProductGroup.Items(1).Text
            pnSelectedProductGroup = ddlProductGroup.Items(1).Value
            ddlProductGroup.Visible = False
        Else
            btnShowAllProducts.Enabled = False
            pnSelectedProductGroup = -1
        End If
    End Sub
    
    Protected Sub ShowProductSelection()
        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        pnlProductList.Visible = True
    End Sub
    
    Protected Sub btnShowAllProducts_Click(ByVal s As Object, ByVal e As EventArgs)
        txtSearchCriteriaAllProducts.Text = ""
        Session("ProductSearchCriteria") = txtSearchCriteriaAllProducts.Text
        lblReportGeneratedDateTime.Visible = True
        Call BindProductGrid("TotalValue")
    End Sub
    
    protected Sub btnGo_Click(ByVal s As Object, ByVal e As EventArgs)
        Session("ProductSearchCriteria") = txtSearchCriteriaAllProducts.Text
        lblReportGeneratedDateTime.visible = True
        Call BindProductGrid("TotalValue")
    End Sub
    
    protected Sub BindProductGrid (SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Report_ProductValues2", oConn)
        Dim sSearchCriteria As String = Session("ProductSearchCriteria")
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
        
        Try
            oAdapter.Fill(oDataSet, "Movements")
            Dim Source As DataView = oDataSet.Tables("Movements").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                dgProducts.Visible = True
                dgProducts.DataSource = Source
                dgProducts.DataBind()
            Else
                dgProducts.Visible = False
                lblError.Text = "no data found"
                lblReportGeneratedDateTime.visible = False
            End If
        Catch ex As SQLException
            lblError.Text = ex.toString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SortProductColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Call BindProductGrid(E.SortExpression)
    End Sub
    
    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
    End Sub
    
    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = -1 Then
            ddlProductGroup.Items.RemoveAt(0)
        End If
        pnSelectedProductGroup = ddl.SelectedValue
        btnShowAllProducts.Enabled = True
    End Sub
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("BHR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("BHR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("BHR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("BHR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_IsProductOwner") = Value
        End Set
    End Property
   
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Product Value Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmProductValueReport" runat="server">
        <asp:Panel id="pnlProductList" runat="server" Width="100%">
            <table id="Table4" runat="server" width="100%">
                <tr>
                    <td valign="Bottom" style="width:0%"></td>
                    <td style="width:50%; white-space:nowrap">
                        <asp:Label ID="Label1" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial">Product Value Report</asp:Label>
                        <br /><br />
                    </td>
                    <td align="right" style="width:50%; white-space:nowrap"></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td style="white-space: nowrap">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr runat="server" id="trProductGroups">
                    <td>
                    </td>
                    <td style="white-space: nowrap">
                        &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged" Visible="False">
                        </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                    <td align="right">
                        <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click" Text="show product groups" Visible="False" /></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td style="white-space: nowrap">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td style="white-space:nowrap">
                <asp:Button ID="btnShowAllProducts" runat="server" Text="show all products" OnClick="btnShowAllProducts_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Label ID="Label2" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana">search:</asp:Label> &nbsp;<asp:TextBox runat="server" Width="100px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtSearchCriteriaAllProducts" MaxLength="50"></asp:TextBox>
                        &nbsp;
                        <asp:Button ID="btnGo" runat="server" Text="go" OnClick="btnGo_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    </td>
                    <td></td>
                </tr>
            </table>
            <br />
            <asp:DataGrid id="dgProducts" runat="server" Width="100%" Font-Names="Arial" Font-Size="XX-Small" CellSpacing="4" AutoGenerateColumns="False" GridLines="None" AllowSorting="True" OnSortCommand="SortProductColumns">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" forecolor="Blue"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticProductKey" SortExpression="LogisticProductKey" HeaderText="No." DataFormatString="{0:000000}">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Code">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDate" SortExpression="ProductDate" HeaderText="Date"></asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription" HeaderText="Description"></asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDepartmentId" SortExpression="ProductDepartmentId" HeaderText="Department">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="LanguageId" SortExpression="LanguageId" HeaderText="Language">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="UnitValue" SortExpression="UnitValue" HeaderText="Unit Value" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Quantity" SortExpression="Quantity" HeaderText="Quantity" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="TotalValue" SortExpression="TotalValue" HeaderText="Total Value" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>&nbsp;
            <asp:Label ID="lblReportGeneratedDateTime" Visible="false" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green"></asp:Label>
        </asp:Panel>
        <br />
        &nbsp;<asp:Label id="lblError" runat="server" font-size="XX-Small" font-names="Arial" forecolor="red"></asp:Label>
    </form>
</body>
</html>
