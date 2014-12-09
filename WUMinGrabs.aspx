<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Const ITEMS_PER_REQUEST As Integer = 30

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Response.RedirectLocation = "http:/my.transworld.eu.com/common/session_expired.aspx"
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
        End If
        Call SetTitle()
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
        Page.Header.Title = sTitle & "Min Grabs Editor"
    End Sub
   
    Protected Sub ddlCustomers_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCustomer.SelectedValue > 1 Then
            Call PopulateMinGrabs()
            trMinGrabs.Visible = True
            Call SetUpdateVisibity(True)
            rcbProduct.Text = String.Empty
            tbMinGrab.Text = String.Empty
            Call PopulateUserGroupDropdown()
        ElseIf ddlCustomer.SelectedValue = 1 Then
            trMinGrabs.Visible = False
            Call SetUpdateVisibity(False)
            Call PopulateMinGrabsAllCustomers()
        Else
            trMinGrabs.Visible = False
            Call SetUpdateVisibity(False)
        End If
        btnSave.Enabled = False
    End Sub

    Protected Function GetEmptyDataText() As String
        If ddlCustomer.SelectedItem.Text.ToLower.Contains("customer") Then
            GetEmptyDataText = "No MIN GRABS defined for any customer"
        Else
            GetEmptyDataText = "No MIN GRABS defined for " & ddlCustomer.SelectedItem.Text
        End If
    End Function
    
    Protected Sub PopulateMinGrabs()
        Dim sSQL As String = "SELECT wumg.id 'MinGrabID', lp.ProductCode + ' - ' + lp.ProductDate + ' (' + lp.ProductDescription + ')' 'Product', upg.GroupName 'UserGroup', MinGrab 'MinGrabValue' FROM ClientData_WU_MinGrabs wumg INNER JOIN LogisticProduct lp ON wumg.LogisticProductKey = lp.LogisticProductKey INNER JOIN UP_UserPermissionGroups upg ON wumg.UserGroup = upg.[id] WHERE lp.CustomerKey = " & ddlCustomer.SelectedValue & " ORDER BY Product"
        Dim dtMinGrabs As DataTable = ExecuteQueryToDataTable(sSQL)
        gvMinGrabs.DataSource = dtMinGrabs
        gvMinGrabs.DataBind()
        trMinGrabs.Visible = True
    End Sub

    Protected Sub PopulateMinGrabsAllCustomers()
        Dim sSQL As String = "SELECT wumg.id 'MinGrabID', lp.ProductCode + ' - ' + lp.ProductDate + ' (' + lp.ProductDescription + ')' 'Product', upg.GroupName 'UserGroup', MinGrab 'MinGrabValue' FROM ClientData_WU_MinGrabs wumg INNER JOIN LogisticProduct lp ON wumg.LogisticProductKey = lp.LogisticProductKey INNER JOIN UP_UserPermissionGroups upg ON wumg.UserGroup = upg.[id] ORDER BY Product"
        Dim dtMinGrabs As DataTable = ExecuteQueryToDataTable(sSQL)
        gvMinGrabs.DataSource = dtMinGrabs
        gvMinGrabs.DataBind()
        trMinGrabs.Visible = True
    End Sub

    Protected Sub PopulateUserGroupDropdown()
        Dim sSQL As String = "SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & ddlCustomer.SelectedValue
        Dim dtUserGroups As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlUserGroup.Items.Clear()
        ddlUserGroup.Items.Add(New ListItem("- please select -", 0))
        For Each drUserGroup As DataRow In dtUserGroups.Rows
            ddlUserGroup.Items.Add(New ListItem(drUserGroup("GroupName"), drUserGroup("id")))
        Next
        
    End Sub
    
    Protected Sub rcbProduct_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        If rcb.SelectedIndex = 0 Then
            btnSave.Enabled = False
        Else
            If ddlUserGroup.SelectedIndex > 0 Then
                btnSave.Enabled = True
            End If
        End If
    End Sub
    
    Protected Sub ddlUserGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            Dim x As String = rcbProduct.SelectedValue
            Dim y As String = rcbProduct.Text
            
            'If IsNumeric(rcbProduct.SelectedValue) Then
            If rcbProduct.Text <> String.Empty Then
                btnSave.Enabled = True
            End If
        Else
            btnSave.Enabled = False
        End If
    End Sub
    
    Protected Function GetProductsByCustomer(ByVal sCustomerKey As String, Optional ByVal sFilter As String = "") As DataTable
        Dim dt As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_WUQuickOrder_GetProducts", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = sCustomerKey
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Filter", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Filter").Value = sFilter

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = 5844   ' marilynfexo
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FavouriteProducts", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@FavouriteProducts").Value = 0

        oAdapter.Fill(dt)
        
        GetProductsByCustomer = dt
    End Function

    Protected Sub rcbProduct_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim s As String = e.Text
        Dim data As DataTable = GetProductsByCustomer(ddlCustomer.SelectedValue, e.Text)
        'Dim sThumbnailImage As String = String.Empty
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbProduct.DataTextField = "Product"
        rcbProduct.DataValueField = "LogisticProductKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            'rcb.Text = data.Rows(i)("Product").ToString() + " (max order: " + data.Rows(i)("Quantity").ToString() + ")"
            rcb.Text = data.Rows(i)("Product").ToString()
            rcb.Value = data.Rows(i)("LogisticProductKey").ToString()
            'sThumbnailImage = data.Rows(i)("ThumbnailImage").ToString()
            rcbProduct.Items.Add(rcb)
            Dim lblProduct As Label = rcb.FindControl("lblProduct")
            Dim imgProduct As Image = rcb.FindControl("imgProduct")
            lblProduct.Text = data.Rows(i)("Product").ToString()
            imgProduct.ImageUrl = "http://my.transworld.eu.com/common/prod_images/thumbs/" & data.Rows(i)("ThumbnailImage").ToString()
        Next
        e.Message = GetStatusMessage(endOffset, data.Rows.Count)
    End Sub

    Private Shared Function GetStatusMessage(ByVal nOffset As Integer, ByVal nTotal As Integer) As String
        If nTotal <= 0 Then
            Return "No matches"
        End If
        If nOffset <= ITEMS_PER_REQUEST Then
            GetStatusMessage = "++++ Click for more items ++++"
        End If
        If nOffset = nTotal Then
            GetStatusMessage = "No more items"
        Else
            GetStatusMessage = "++++ Click for more items ++++"
        End If
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


    Protected Sub lnkbtnHide_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trMinGrabs.Visible = False
    End Sub
    
    Protected Sub lnkbtnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim MinGrabID As Int32 = lnkbtn.CommandArgument
        Dim sSQL As String = "DELETE FROM ClientData_WU_MinGrabs WHERE [id] = " & MinGrabID
        Call ExecuteQueryToDataTable(sSQL)
        Call PopulateMinGrabs()
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbMinGrab.Text) Then
            WebMsgBox.Show("Please enter a valid number for the MIN GRAB quantity.")
            Exit Sub
        End If
        Call SaveMinGrab()
    End Sub
    
    Protected Sub SaveMinGrab()
        Dim sSQL As String = "DELETE FROM ClientData_WU_MinGrabs WHERE LogisticProductKey = " & rcbProduct.SelectedValue & " AND UserGroup = " & ddlUserGroup.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
        sSQL = "INSERT INTO ClientData_WU_MinGrabs (LogisticProductKey, UserGroup, MinGrab) VALUES (" & rcbProduct.SelectedValue & ", " & ddlUserGroup.SelectedValue & ", " & CInt(tbMinGrab.Text).ToString & ")"
        Call ExecuteQueryToDataTable(sSQL)
        Call PopulateMinGrabs()
    End Sub

    Protected Sub SetUpdateVisibity(ByVal bVisible As Boolean)
        trUpdate01.Visible = bVisible
        trUpdate02.Visible = bVisible
        trUpdate03.Visible = bVisible
        trUpdate04.Visible = bVisible
        trUpdate05.Visible = bVisible
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server" />
    <asp:Label ID="lblLegendCustomer0" runat="server" Font-Names="Verdana" Font-Size="Small" Font-Bold="true" Text="MIN GRABS Editor" />
    &nbsp;<asp:Panel ID="pnlUpdate" Width="100%" runat="server">
        &nbsp;<table style="width: 100%">
            <tr>
                <td style="width: 10%">
                </td>
                <td style="width: 25%">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblLegendCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlCustomer" runat="server" AutoPostBack="True" Font-Size="X-Small" OnSelectedIndexChanged="ddlCustomers_SelectedIndexChanged">
                        <asp:ListItem Value="0" Selected="True">- please select -</asp:ListItem>
                        <asp:ListItem Value="1">- ALL CUSTOMERS -</asp:ListItem>
                        <asp:ListItem Value="579">WURS</asp:ListItem>
                        <asp:ListItem Value="686">WUIRE</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trMinGrabs" runat="server" visible="false">
                <td>
                    <asp:Label ID="lblLegendMinGrabs" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Min Grabs:"></asp:Label>
                    <br />
                    &nbsp;<asp:LinkButton ID="lnkbtnHide" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnHide_Click">hide</asp:LinkButton>
                </td>
                <td colspan="2">
                    <asp:GridView ID="gvMinGrabs" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="90%" AutoGenerateColumns="False">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnRemove" runat="server" CommandArgument='<%# Container.DataItem("MinGrabID")%>' OnClick="lnkbtnRemove_Click">remove</asp:LinkButton>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Product" HeaderText="Product" ReadOnly="True" SortExpression="Product" />
                            <asp:BoundField DataField="UserGroup" HeaderText="User Group" ReadOnly="True" SortExpression="UserGroup" />
                            <asp:BoundField DataField="MinGrabValue" HeaderText="Min Grab Value" ReadOnly="True" SortExpression="MinGrabValue" />
                        </Columns>
                        <EmptyDataTemplate>
                            <asp:Label ID="lblLegendEmptyDataMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text='<%# GetEmptyDataText() %>' />
                        </EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <hr />
                </td>
            </tr>
            <tr id="trUpdate01" runat="server" visible="false">
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblLegendDefine" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="DEFINE A NEW MIN GRAB OR EDIT AN EXISTING MIN GRAB" Font-Bold="True" />
                </td>
            </tr>
            <tr id="trUpdate02" runat="server" visible="false">
                <td>
                    <asp:Label ID="lblLegendProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Product:"></asp:Label>
                </td>
                <td colspan="2">
                    <telerik:RadComboBox ID="rcbProduct" runat="server" AutoPostBack="True" CausesValidation="False" EmptyMessage="Click here to show products, or type a search phrase" EnableLoadOnDemand="True" EnableVirtualScrolling="True" Filter="Contains" Font-Bold="true" Font-Names="Arial" Font-Size="X-Small" HighlightTemplatedItems="true" OnItemsRequested="rcbProduct_ItemsRequested" OnSelectedIndexChanged="rcbProduct_SelectedIndexChanged" ShowMoreResultsBox="True" ToolTip="Shows all available products when no search text is specified. Search for products by typing a product code or description." Width="100%">
                        <ItemTemplate>
                            <table>
                                <tr>
                                    <td style="width: 70px">
                                        <asp:Image ID="imgProduct" runat="server" />
                                    </td>
                                    <td style="width: 220px">
                                        <asp:Label ID="lblProduct" runat="server" />
                                    </td>
                                </tr>
                            </table>
                        </ItemTemplate>
                    </telerik:RadComboBox>
                </td>
            </tr>
            <tr id="trUpdate03" runat="server" visible="false">
                <td>
                    <asp:Label ID="lblLegendUserGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="User group:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlUserGroup" runat="server" AutoPostBack="True" Font-Size="X-Small" OnSelectedIndexChanged="ddlUserGroup_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trUpdate04" runat="server" visible="false">
                <td>
                    <asp:Label ID="lblLegendMinGrab" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Min grab:" />
                </td>
                <td>
                    <asp:TextBox ID="tbMinGrab" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="60px" />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trUpdate05" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Button ID="btnSave" runat="server" Text="Save" Width="200px" OnClick="btnSave_Click" Enabled="False" />
                </td>
                <td>
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
