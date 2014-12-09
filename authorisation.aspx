<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
    ' Selected Product
    ' Other filtering
    ' Authorisable Products
    
    Dim gsConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call HideAllPanels()
            lblSelectedCustomer.Text = String.Empty
        End If
        btnCustomersWithAuthorisableProducts.Focus()
        'tbCustomer.Focus()
        'tbCustomer.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
    End Sub

    Protected Sub btnCustomersWithAuthorisableProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetCustomersWithAuthorisableProducts()
    End Sub

    Protected Sub GetCustomersWithAuthorisableProducts()
        Dim sSQL As String = "SELECT DISTINCT CustomerAccountCode + ' (' + CAST(c.CustomerKey AS varchar(4)) + ')', c.CustomerKey FROM LogisticProductAuthorisable lpa INNER JOIN LogisticProduct lp ON lpa.LogisticProductKey = lp.LogisticProductKey INNER JOIN Customer c ON c.CustomerKey = lp.CustomerKey WHERE CustomerStatusId = 'ACTIVE'"
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        lbCustomersWithAuthorisableProducts.Items.Clear()
        Dim li As New ListItem
        li.Text = "all customers"
        li.Value = 0
        lbCustomersWithAuthorisableProducts.Items.Add(li)
        For Each dr As DataRow In oDataTable.Rows
            Dim liEntry As New ListItem
            liEntry.Text = dr(0)
            liEntry.Value = dr(1)
            lbCustomersWithAuthorisableProducts.Items.Add(liEntry)
        Next
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub btnCustomersWithAuthorisations_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetCustomersWithAuthorisations()
    End Sub

    Protected Sub GetCustomersWithAuthorisations()
        Dim sSQL As String = "SELECT DISTINCT CustomerAccountCode + ' (' + CAST(c.CustomerKey AS varchar(4)) + ')', c.CustomerKey FROM LogisticProductAuthorisation lpa INNER JOIN LogisticProduct lp ON lpa.LogisticProductKey = lp.LogisticProductKey INNER JOIN Customer c ON c.CustomerKey = lp.CustomerKey WHERE CustomerStatusId = 'ACTIVE' "
        If rblAuthorisationType.SelectedIndex = 0 Then
            sSQL += " AND AuthorisationExpiryDateTime >= GETDATE()"
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        lbCustomersWithAuthorisations.Items.Clear()
        Dim li As New ListItem
        li.Text = "all customers"
        li.Value = 0
        lbCustomersWithAuthorisations.Items.Add(li)
        For Each dr As DataRow In oDataTable.Rows
            Dim liEntry As New ListItem
            liEntry.Text = dr(0)
            liEntry.Value = dr(1)
            lbCustomersWithAuthorisations.Items.Add(liEntry)
        Next
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub HideAllPanels()
        pnlData.Visible = False
    End Sub

    Protected Sub lbCustomersWithAuthorisableProducts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        psSelectedCustomerKey = lb.SelectedValue
        Call PopulateClientAuthorisations(lb.SelectedValue)
        Call InitProductDropdown(lb.SelectedValue)
        pnlData.Visible = True
        lblSelectedCustomer.Text = lb.SelectedItem.Text
    End Sub

    Protected Sub lbCustomersWithAuthorisations_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        psSelectedCustomerKey = lb.SelectedValue
        Call PopulateClientAuthorisations(lb.SelectedValue)
        Call InitProductDropdown(lb.SelectedValue)
        pnlData.Visible = True
        lblSelectedCustomer.Text = lb.SelectedItem.Text
    End Sub
    
    Protected Sub InitProductDropdown(ByVal sCustomerKey As String)
        Dim sSQL As String = String.Empty
        sSQL += "SELECT DISTINCT lpa1.LogisticProductKey, ProductCode, '(' + CAST(lpa1.LogisticProductKey AS varchar(6)) + ') ' + ProductCode + ', ' + ProductDate + ', ' + ProductDescription "
        sSQL += "FROM LogisticProductAuthorisation lpa1 "
        sSQL += "LEFT OUTER JOIN LogisticProduct lp "
        sSQL += "ON lp.LogisticProductKey = lpa1.LogisticProductKey "
        If Not sCustomerKey = "0" Then
            sSQL += "WHERE lpa1.CustomerKey = " & sCustomerKey
        End If
        sSQL += "ORDER BY ProductCode "
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        Dim li As New ListItem
        ddlProduct.Items.Clear()
        li.Value = 0
        li.Text = "(all products)"
        ddlProduct.Items.Add(li)
        For Each dr As DataRow In oDataTable.Rows
            Dim liEntry As New ListItem
            If IsDBNull(dr(0)) Then
                liEntry.Value = 0
            Else
                liEntry.Value = dr(0)
            End If
            If IsDBNull(dr(2)) Then
                liEntry.Text = "(null)"
            Else
                liEntry.Text = dr(2)
            End If
            ddlProduct.Items.Add(liEntry)
        Next
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub PopulateClientAuthorisations(ByVal sCustomerKey As String)
        Dim sSQL As String = String.Empty
        sSQL += "SELECT '(' + CAST(lpa1.LogisticProductKey AS varchar(6)) + ') ' + lp.ProductCode 'Product Code', "
        sSQL += "lp.ProductDate 'Version/Date', "
        sSQL += "lp.ProductDescription 'Description', "
        sSQL += "'(' + CAST(lpa1.UserProfileKey AS varchar(6)) + ') ' + up2.FirstName + ' ' + up2.LastName + ' (' + up2.UserId + ')' 'Requested by', "
        sSQL += "'(' + CAST(lpa2.Authoriser AS varchar(6)) + ') ' + up1.FirstName + ' ' + up1.LastName + ' (' + up1.UserId + ')' 'Authoriser', "
        sSQL += "AuthorisationGUID 'GUID', "
        sSQL += "AuthorisationRequestDateTime 'Requested', "
        sSQL += "AuthorisationGrantDateTime 'Granted', "
        sSQL += "AuthorisationExpiryDateTime 'Expires', "
        sSQL += "AuthorisedQuantity 'Authd Qty', "
        sSQL += "QuantityRemaining 'Qty Remaining', "
        sSQL += "Granted 'Granted' "
        sSQL += "FROM LogisticProduct lp "
        sSQL += "LEFT OUTER JOIN LogisticProductAuthorisation lpa1 "
        sSQL += "ON lpa1.LogisticProductKey = lp.LogisticProductKey "
        sSQL += "LEFT OUTER JOIN LogisticProductAuthorisable lpa2 "
        sSQL += "ON lpa2.LogisticProductKey = lpa1.LogisticProductKey "
        sSQL += "LEFT OUTER JOIN UserProfile up1 "
        sSQL += "ON lpa2.Authoriser = up1.[Key] "
        sSQL += "LEFT OUTER JOIN UserProfile up2 "
        sSQL += "ON lpa1.UserProfileKey = up2.[Key] "
        sSQL += "WHERE lpa1.LogisticProductKey IN "
        sSQL += "(SELECT LogisticProductKey "
        sSQL += "FROM LogisticProductAuthorisation lpa1 "
        If sCustomerKey = "0" Then
            sSQL += ")"
        Else
            sSQL += "WHERE lpa1.CustomerKey = " & sCustomerKey & ")"
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvAuthorisations.DataSource = oDataTable
        gvAuthorisations.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub PopulateClientAuthorisationsByProduct(ByVal sLogisticProductKey As String)
        Dim sSQL As String = String.Empty
        sSQL += "SELECT '(' + CAST(lpa1.LogisticProductKey AS varchar(6)) + ') ' + lp.ProductCode 'Product Code', "
        sSQL += "lp.ProductDate 'Version/Date', "
        sSQL += "lp.ProductDescription 'Description', "
        sSQL += "'(' + CAST(lpa1.UserProfileKey AS varchar(6)) + ') ' + up2.FirstName + ' ' + up2.LastName + ' (' + up2.UserId + ')' 'Requested by', "
        sSQL += "'(' + CAST(lpa2.Authoriser AS varchar(6)) + ') ' + up1.FirstName + ' ' + up1.LastName + ' (' + up1.UserId + ')' 'Authoriser', "
        sSQL += "AuthorisationGUID 'GUID', "
        sSQL += "AuthorisationRequestDateTime 'Requested', "
        sSQL += "AuthorisationGrantDateTime 'Granted', "
        sSQL += "AuthorisationExpiryDateTime 'Expires', "
        sSQL += "AuthorisedQuantity 'Authd Qty', "
        sSQL += "QuantityRemaining 'Qty Remaining', "
        sSQL += "Granted 'Granted' "
        sSQL += "FROM LogisticProduct lp "
        sSQL += "LEFT OUTER JOIN LogisticProductAuthorisation lpa1 "
        sSQL += "ON lpa1.LogisticProductKey = lp.LogisticProductKey "
        sSQL += "LEFT OUTER JOIN LogisticProductAuthorisable lpa2 "
        sSQL += "ON lpa2.LogisticProductKey = lpa1.LogisticProductKey "
        sSQL += "LEFT OUTER JOIN UserProfile up1 "
        sSQL += "ON lpa2.Authoriser = up1.[Key] "
        sSQL += "LEFT OUTER JOIN UserProfile up2 "
        sSQL += "ON lpa1.UserProfileKey = up2.[Key] "
        sSQL += "WHERE lp.LogisticProductKey = " & sLogisticProductKey
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvAuthorisations.DataSource = oDataTable
        gvAuthorisations.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Property psSelectedCustomerKey() As String
        Get
            Dim o As Object = ViewState("PA_SelectedCustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PA_SelectedCustomerKey") = Value
        End Set
    End Property
    
    Property psSelectedProduct() As String
        Get
            Dim o As Object = ViewState("PA_SelectedProduct")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PA_SelectedProduct") = Value
        End Set
    End Property
    
    Property psCustomerSortExpression() As String
        Get
            Dim o As Object = ViewState("CU_CustomerSortExpression")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerSortExpression") = Value
        End Set
    End Property
    
    Protected Sub ddlProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        psSelectedProduct = ddl.SelectedValue
        If psSelectedProduct = "0" Then
            Call PopulateClientAuthorisations(psSelectedCustomerKey)
        Else
            Call PopulateClientAuthorisationsByProduct(psSelectedProduct)
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml ">
<head id="Head1" runat="server">
    <title>Authorisation Info</title>
    <link href="Reports.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <strong>Product Authorisation<br />
            </strong>
            <br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 24%" align="right">
                    </td>
                    <td style="width: 24%" align="left">
                    </td>
                    <td style="width: 4%">
                    </td>
                    <td style="width: 24%" align="right">
                    </td>
                    <td style="width: 24%" align="left">
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        Find customers with one or more authorisable products:<br />
                        <br />
                        <asp:Button ID="btnCustomersWithAuthorisableProducts" runat="server" Text="go" OnClick="btnCustomersWithAuthorisableProducts_Click" /></td>
                    <td>
                        &nbsp;<asp:ListBox ID="lbCustomersWithAuthorisableProducts" runat="server" Rows="6"
                            Width="200px" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="lbCustomersWithAuthorisableProducts_SelectedIndexChanged"
                            AutoPostBack="True"></asp:ListBox></td>
                    <td>
                        &nbsp;</td>
                    <td align="center">
                        Find customers with one or more authorisations:<br />
                        <br />
                        <asp:Button ID="btnCustomersWithAuthorisations" runat="server" Text="go" OnClick="btnCustomersWithAuthorisations_Click" /><br />
                        <br />
                        <asp:RadioButtonList ID="rblAuthorisationType" runat="server">
                            <asp:ListItem Selected="True">unexpired authorisations</asp:ListItem>
                            <asp:ListItem Value="0">any authorisation</asp:ListItem>
                        </asp:RadioButtonList></td>
                    <td>
                        <asp:ListBox ID="lbCustomersWithAuthorisations" runat="server" Rows="6" Width="200px"
                            Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="lbCustomersWithAuthorisations_SelectedIndexChanged"
                            AutoPostBack="True"></asp:ListBox></td>
                </tr>
            </table>
        </div>
        <br />
        <asp:Panel ID="pnlData" runat="server" Width="100%">
            <strong>Data for </strong>
            <asp:Label ID="lblSelectedCustomer" runat="server"></asp:Label>
            &nbsp; &nbsp; &nbsp; &nbsp;
            <asp:Label ID="Label1" runat="server" Text="select a product: "></asp:Label>
            <asp:DropDownList ID="ddlProduct" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlProduct_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small">
            </asp:DropDownList>
            <br />
            <strong>
                <br />
                Authorisations
                <br />
            </strong>
            <br />
            <asp:GridView ID="gvAuthorisations" runat="server" Width="100%" CellPadding="2" Font-Names="Verdana"
                Font-Size="XX-Small">
            </asp:GridView>
            &nbsp;<span style="color: #ff0000"></span><br />
            <strong>Authorisable Products&nbsp;</strong><br />
            &nbsp;
            <br />
            <br />
            <asp:GridView ID="gvAuthorisableProducts" runat="server" Width="100%" CellPadding="2"
                Font-Names="Verdana" Font-Size="XX-Small">
            </asp:GridView>
        </asp:Panel>
    </form>

    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>

</body>
</html>