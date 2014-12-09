<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<script runat="server">

    ' put in warning of 'hidden' virtual products by retrieving count of virtual products with no restriction on DELETED or ARCHIVED and comparing it with displayed value
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            ddlCustomer.Focus()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Virtual Product Manager"
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
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

    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = 0 Then
            gvPhysicalProducts.DataSource = Nothing
            gvPhysicalProducts.DataBind()
            gvVirtualProducts.DataSource = Nothing
            gvVirtualProducts.DataBind()
            pnlControls.Visible = False
        Else
            'Dim sSQL As String = "SELECT LogisticProductKey, ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND ArchiveFlag = 'N' AND DeletedFlag = 'N' AND NOT LogisticProductKey IN (SELECT DISTINCT VirtualProductKey FROM LogisticVirtualProduct) ORDER BY ProductCode, ProductDate"
            'Dim dtProducts As DataTable = ExecuteQueryToDataTable(sSQL)
            'gvPhysicalProducts.DataSource = dtProducts
            'gvPhysicalProducts.DataBind()
            Call BindPhysicalProductsGridView()
            Call GetVirtualProductsForCustomer()
            Call PopulateVirtualProductsDropdown()
            pnlControls.Visible = True
        End If
        pnlEditVirtualProduct.Visible = False
    End Sub
    
    Protected Sub BindPhysicalProductsGridView()
        Dim sSQL As String = "SELECT LogisticProductKey, ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND ArchiveFlag = 'N' AND DeletedFlag = 'N' AND NOT LogisticProductKey IN (SELECT DISTINCT VirtualProductKey FROM LogisticVirtualProduct) ORDER BY ProductCode, ProductDate"
        Dim dtProducts As DataTable = ExecuteQueryToDataTable(sSQL)
        gvPhysicalProducts.DataSource = dtProducts
        gvPhysicalProducts.DataBind()
    End Sub
    
    Protected Sub GetVirtualProductsForCustomer()
        Dim nCustomerKey As Int32 = ddlCustomer.SelectedValue
        Dim sSQL As String = "SELECT DISTINCT VirtualProductKey, lp.LogisticProductKey, lp.ProductCode + ' - ' + ISNULL(lp.ProductDate, '') + ' ' + lp.ProductDescription 'Product', '0' 'Contents' FROM LogisticVirtualProduct lvp INNER JOIN LogisticProduct lp ON lvp.VirtualProductKey = lp.LogisticProductKey WHERE lp.CustomerKey = " & nCustomerKey & " AND lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' ORDER BY Product"
        Dim dtVirtualProducts As DataTable = ExecuteQueryToDataTable(sSQL)
        gvVirtualProducts.DataSource = dtVirtualProducts
        gvVirtualProducts.DataBind()
    End Sub
    
    Protected Sub PopulateVirtualProductsDropdown()
        ddlVirtualProduct.Items.Clear()
        Dim sSQL As String = "SELECT ProductCode + ' - ' + ISNULL(ProductDate, '') + ' ' + ProductDescription 'Product', LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND DeletedFlag = 'N' AND ArchiveFlag = 'N' AND NOT LogisticProductKey IN (SELECT VirtualProductKey FROM LogisticVirtualProduct) ORDER BY ProductCode"
        Dim dtProducts As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlVirtualProduct.Items.Add(New ListItem("- please select -", 0))
        For Each drProduct In dtProducts.Rows
            ddlVirtualProduct.Items.Add(New ListItem(drProduct("Product"), drProduct("LogisticProductKey")))
        Next
    End Sub
    
    Protected Sub gvVirtualProducts_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidVirtualProductKey As HiddenField
            hidVirtualProductKey = gvr.Cells(0).FindControl("hidVirtualProductKey")
            Dim sSQL As String = "SELECT lp.ProductCode + ' - ' + ISNULL(lp.ProductDate, '') + ' ' + lp.ProductDescription 'Product', Qty FROM LogisticVirtualProduct lvp INNER JOIN LogisticProduct lp ON lvp.LogisticProductKey = lp.LogisticProductKey WHERE VirtualProductKey = " & hidVirtualProductKey.Value
            Dim dtPhysicalProducts As DataTable = ExecuteQueryToDataTable(sSQL)
            Dim sbPhysicalProducts As New StringBuilder
            For Each drPhysicalProduct As DataRow In dtPhysicalProducts.Rows
                sbPhysicalProducts.Append(drPhysicalProduct(0))
                sbPhysicalProducts.Append(" ")
                sbPhysicalProducts.Append(drPhysicalProduct(1))
                sbPhysicalProducts.Append("<br />")
            Next
            gvr.Cells(2).Text = sbPhysicalProducts.ToString
        End If
    End Sub

    Protected Sub lnkbtnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nVirtualProductKey As Int32 = lnkbtn.CommandArgument
        Dim sSQL As String = "DELETE FROM LogisticVirtualProduct WHERE VirtualProductKey = " & nVirtualProductKey
        Call ExecuteQueryToDataTable(sSQL)
        Call GetVirtualProductsForCustomer()
        Call PopulateVirtualProductsDropdown()
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bValid As Boolean = True
        Dim nProductCount As Int32 = 0
        Dim sSQL As String
        For Each gvr As GridViewRow In gvPhysicalProducts.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim cbEnabled As CheckBox = gvr.Cells(0).FindControl("cbEnabled")
                If cbEnabled.Checked Then
                    Dim tbQty As TextBox = gvr.Cells(1).FindControl("tbQty")
                    If IsNumeric(tbQty.Text) AndAlso CInt(tbQty.Text) > 0 Then
                        nProductCount += 1
                    Else
                        WebMsgBox.Show("Non-numeric or negative quantity found - please correct this.")
                        Exit Sub
                    End If
                    
                End If
            End If
        Next
        If nProductCount > 0 Then
            sSQL = "DELETE FROM LogisticVirtualProduct WHERE VirtualProductKey = " & ddlVirtualProduct.SelectedValue
            Call ExecuteQueryToDataTable(sSQL)
            For Each gvr As GridViewRow In gvPhysicalProducts.Rows
                If gvr.RowType = DataControlRowType.DataRow Then
                    Dim cbEnabled As CheckBox = gvr.Cells(0).FindControl("cbEnabled")
                    If cbEnabled.Checked Then
                        Dim tbQty As TextBox = gvr.Cells(1).FindControl("tbQty")
                        Dim hidLogisticProductKey As HiddenField = gvr.Cells(0).FindControl("hidLogisticProductKey")
                        sSQL = "INSERT INTO LogisticVirtualProduct (VirtualProductKey, LogisticProductKey, Qty, LastUpdatedOn, LastUpdatedBy) VALUES (" & ddlVirtualProduct.SelectedValue & ", " & hidLogisticProductKey.Value & ", " & CInt(tbQty.Text).ToString & ", GETDATE(), 0)"
                        Call ExecuteQueryToDataTable(sSQL)
                    End If
                End If
            Next
            Call GetVirtualProductsForCustomer()
            ddlVirtualProduct.SelectedIndex = 0
            pnlEditVirtualProduct.Visible = False
        Else
            WebMsgBox.Show("You must have at least one physical product defined for a virtual product.")
        End If
    End Sub
    
    Protected Sub ddlVirtualProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex = 0 Then
            pnlEditVirtualProduct.Visible = False
        Else
            pnlEditVirtualProduct.Visible = True
            Call BindPhysicalProductsGridView()
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" runat="Server">
    <main:Header id="ctlHeader" runat="server"/>
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <asp:Label ID="lblLegendVirtualProductManager" runat="server" Font-Names="Verdana" Font-Size="X-Small" Text="Virtual Product Manager" Font-Bold="True" />
    <br />
    <asp:Label ID="lblLegendCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer:" />
    <asp:DropDownList ID="ddlCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged" AutoPostBack="True">
        <asp:ListItem Selected="True" Value="0">- please select -</asp:ListItem>
        <asp:ListItem Value="579">WURS</asp:ListItem>
        <asp:ListItem Value="686">WUIRE</asp:ListItem>
    </asp:DropDownList>
    <br />
    <asp:Panel ID="pnlControls" runat="server" Width="100%" Visible="false">
        <br />
        <asp:Label ID="lblLegendVirtualProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Virtual Products:" />
        <br />
        <asp:GridView ID="gvVirtualProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" CellPadding="2" OnRowDataBound="gvVirtualProducts_RowDataBound">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnRemove" runat="server" CommandArgument='<%# Container.DataItem("VirtualProductKey")%>' Text="remove" OnClick="lnkbtnRemove_Click" />
                        <asp:HiddenField ID="hidVirtualProductKey" runat="server" Value='<%# Container.DataItem("VirtualProductKey")%>' />
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" Width="80px" />
                </asp:TemplateField>
                <asp:BoundField DataField="Product" HeaderText="Virtual Product" ReadOnly="True" SortExpression="ProductCode" />
                <asp:BoundField DataField="Contents" HeaderText="Physical Products" ReadOnly="True" SortExpression="Contents" />
            </Columns>
            <EmptyDataTemplate>
                no virtual products found
            </EmptyDataTemplate>
        </asp:GridView>
        <br />
        <hr />
        <asp:Label ID="lblLegendHeader" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="true" Text="CREATE OR EDIT A VIRTUAL PRODUCT" />
        <br />
        <br />
        <asp:Label ID="lblLegendVirtualProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Virtual Product:" />
        <asp:DropDownList ID="ddlVirtualProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlVirtualProduct_SelectedIndexChanged" />
        <br />
        <asp:Panel ID="pnlEditVirtualProduct" runat="server" Width="100%" Visible="false">
            <br />
            <asp:Label ID="lblLegendPhysicalProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Physical Products:" />
            <br />
            <asp:GridView ID="gvPhysicalProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" CellPadding="2">
                <Columns>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:CheckBox ID="cbEnabled" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                            <asp:HiddenField ID="hidLogisticProductKey" runat="server" Value='<%# Container.DataItem("LogisticProductKey")%>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" Width="40px" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Quantity">
                        <ItemTemplate>
                            <asp:TextBox ID="tbQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="50px" />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Center" Width="80px" />
                    </asp:TemplateField>
                    <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductCode" />
                    <asp:BoundField DataField="ProductDate" HeaderText="Value / Date" ReadOnly="True" SortExpression="ProductDate" />
                    <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                </Columns>
            </asp:GridView>
            <asp:Button ID="btnSave" runat="server" OnClick="btnSave_Click" Text="Save" Width="200px" />
            <br />
        </asp:Panel>
    </asp:Panel>
    <asp:Panel ID="pnlHelp" runat="server" Font-Names="Verdana" Font-Size="X-Small" Width="100%">
        <br />
        <hr />
        <strong>HOW TO CREATE A VIRTUAL PRODUCT<br />
        </strong>
        <br />
        1. Use Product Manager to create the product you want as your virtual product, for the relevant customer.<br />
        <br />
        2.&nbsp; Use the AIMS Desktop application to assign this product some stock quantity, placing the quanity in the DEMO warehouse.<br />
        <br />
        3.&nbsp; Using the Virtual Product Manager (this tab) select the relvant customer, then select the product you have just created.&nbsp; Assign the required physical products with the check box and quantity field, then click the Save button.<br />
        <br />
        <strong>VIRTUAL PRODUCT BEHAVIOUR WHEN ORDERING<br />
        <br />
        </strong>When you select a virtual product on the Quick Order tab, the physical products that make up the virtual product will be placed in the basket, if the max grab and product credit amounts permit. If one or more physical products fails to meet the max grab or product credit amount check, none of the products will be added.<br />
        <br />
        <strong>NOTES</strong>
        <br />
        <br />
        You can assign a physical product to more than one virtual product.<br />
        <br />
        You must not assign a virtual product to be part of another virtual product (but the system won&#39;t stop you doing that).<br />
        <br />
        A virtual product cannot be part of itself (but the system won&#39;t stop you doing that, either).<br />
        <br />
        Virtual products are recognised by the Quick Order tab, but <strong>NOT</strong> by the classic order tab.<br />
        <br />
        [end]<br />
    </asp:Panel>
    <br />
    </form>
</body>
</html>
