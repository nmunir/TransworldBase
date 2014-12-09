<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    Const CUSTOMER_BLACKROCK As Int32 = 23
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Call SetTitle()
            Call HideAllPanelsAndRows()
            Call PopulateUserDropdown()
            Call RefreshAlertsGridView()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "BlackRock Low Stock Alerts"
    End Sub
   
    Protected Sub HideAllPanelsAndRows()
    End Sub
    
    Protected Sub CheckForOrphanCostCentres()
        
    End Sub
    
    Protected Sub PopulateUserDropdown()
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' 'Username', [key] 'UserKey' FROM UserProfile WHERE DeletedFlag = 0 AND CustomerKey = " & CUSTOMER_BLACKROCK & " ORDER BY LastName"
        Dim dtUsers As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlUser.Items.Clear()
        ddlUser.Items.Add(New ListItem("- please select -", 0))
        For Each dtUser As DataRow In dtUsers.Rows
            ddlUser.Items.Add(New ListItem(dtUser("Username"), dtUser("UserKey")))
        Next
    End Sub

    Protected Sub PopulateCostCentreListbox()
        Dim sSQL As String = "SELECT DISTINCT ISNULL(ProductDepartmentId, '') 'ProductDepartmentId' FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_BLACKROCK & " AND DeletedFlag = 'N' AND ProductDepartmentId IS NOT NULL AND ProductDepartmentId <> '' ORDER BY ProductDepartmentId"
        Dim dtCostCentres As DataTable = ExecuteQueryToDataTable(sSQL)
        lbCostCentre.Items.Clear()
        For Each dtCostCentre As DataRow In dtCostCentres.Rows
            lbCostCentre.Items.Add(New ListItem(dtCostCentre("ProductDepartmentId"), dtCostCentre("ProductDepartmentId")))
        Next
        lbCostCentre.Rows = dtCostCentres.Rows.Count
    End Sub

    Protected Sub btnUpdate_Click(sender As Object, e As System.EventArgs)
        Call Update()
    End Sub
    
    Protected Sub Update()
        Dim sSQL As String
        Call DeleteCostCentresForUser()
        For i = 0 To lbCostCentre.Items.Count - 1
            If lbCostCentre.Items(i).Selected Then
                sSQL = "INSERT INTO StockAlertCostCentreRecipients (UserKey, CostCentre, LastUpdatedOn, LastUpdatedBy) VALUES (" & ddlUser.SelectedValue & ", '" & lbCostCentre.Items(i).Text.Replace("'", "''") & "', GETDATE(), 0)"
                Call ExecuteQueryToDataTable(sSQL)
            End If
        Next
        Call RefreshAlertsGridView()
    End Sub

    Protected Sub RefreshAlertsGridView()
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' 'Username', [key] 'UserKey' FROM UserProfile WHERE DeletedFlag = 0 AND CustomerKey = " & CUSTOMER_BLACKROCK & " ORDER BY LastName"
        Dim dtUsers As DataTable = ExecuteQueryToDataTable(sSQL)
        gvCostCentresByUser.DataSource = dtUsers
        gvCostCentresByUser.DataBind()
        
        sSQL = "SELECT DISTINCT CostCentre FROM StockAlertCostCentreRecipients ORDER BY CostCentre"
        Dim dtCostCentres As DataTable = ExecuteQueryToDataTable(sSQL)
        gvUsersByCostCentre.DataSource = dtCostCentres
        gvUsersByCostCentre.DataBind()
    End Sub
    
    Protected Sub DeleteCostCentresForUser()
        Dim sSQL As String = "DELETE FROM StockAlertCostCentreRecipients WHERE UserKey = " & ddlUser.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub ddlUser_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call InitUserCostCentres(ddl.SelectedValue)
        If ddl.Items(0).Value = 0 Then
            ddl.Items.RemoveAt(0)
            Call PopulateCostCentreListbox()
        End If
        btnUpdate.Enabled = True
    End Sub
    
    Protected Sub InitUserCostCentres(nUserKey As Int32)
        Dim sSQL As String = "SELECT CostCentre FROM StockAlertCostCentreRecipients WHERE UserKey = " & nUserKey
        Dim dtCostCentres As DataTable = ExecuteQueryToDataTable(sSQL)
        Call CostCentreListboxDeselectAll()
        For Each drCostCentre As DataRow In dtCostCentres.Rows
            Dim sCostCentre As String = drCostCentre(0)
            For i = 0 To lbCostCentre.Items.Count - 1
                If lbCostCentre.Items(i).Text = sCostCentre Then
                    lbCostCentre.Items(i).Selected = True
                End If
            Next
        Next
    End Sub
    
    Protected Sub CostCentreListboxDeselectAll()
        For i = 0 To lbCostCentre.Items.Count - 1
            lbCostCentre.Items(i).Selected = False
        Next
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

    Protected Sub gvCostCentresByUser_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            'Dim lbl As Label = gvr.Cells(1).FindControl("lblCostCentres")
            Dim hidUserKey As HiddenField = gvr.FindControl("hidUserKey")
            Dim lbl As Label = gvr.FindControl("lblCostCentres")
            'lbl.Text = "FOUND IT! " & hidUserKey.Value
            Dim sSQL As String = "SELECT CostCentre FROM StockAlertCostCentreRecipients WHERE UserKey = " & hidUserKey.Value
            Dim dtCostCentres As DataTable = ExecuteQueryToDataTable(sSQL)
            For Each drCostCentre As DataRow In dtCostCentres.Rows
                Dim sCostCentre As String = drCostCentre(0)
                lbl.Text &= sCostCentre & ", "
            Next
            If lbl.Text.Length > 0 Then
                lbl.Text = lbl.Text.Substring(0, lbl.Text.Length - 2)
            End If
        End If
    End Sub
    
    Protected Sub gvUsersByCostCentre_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim lblCostCentre As Label = gvr.FindControl("lblCostCentre")
            Dim lblUsers As Label = gvr.FindControl("lblUsers")
            Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' 'Username' FROM StockAlertCostCentreRecipients saccr INNER JOIN UserProfile up ON saccr.UserKey = up.[key] WHERE CostCentre = '" & lblCostCentre.Text.Replace("'", "''") & "'"
            Dim dtUsers As DataTable = ExecuteQueryToDataTable(sSQL)
            For Each drUser As DataRow In dtUsers.Rows
                lblUsers.Text &= drUser(0) & ", "
            Next
            If lblUsers.Text.Length > 0 Then
                lblUsers.Text = lblUsers.Text.Substring(0, lblUsers.Text.Length - 2)
            End If
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <%--<main:Header ID="ctlHeader" runat="server" />--%>
    <table style="width: 100%">
        <tr>
            <td style="width: 2%">
                &nbsp;
            </td>
            <td style="width: 26%">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Gray">BlackRock Low Stock Alerts</asp:Label>
            </td>
            <td style="width: 40%">
                &nbsp;
            </td>
            <td style="width: 30%">
                &nbsp;
            </td>
            <td style="width: 2%">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td valign="top">
                <asp:Label ID="lblLegendUser" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small">User:</asp:Label>
                <br />
                <asp:DropDownList ID="ddlUser" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small" AutoPostBack="True" 
                    onselectedindexchanged="ddlUser_SelectedIndexChanged">
                </asp:DropDownList>
            </td>
            <td>
                <asp:Label ID="lblLegendCostCentre" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small">Cost Centre:</asp:Label>
                <asp:ListBox ID="lbCostCentre" runat="server" Rows="20" 
                    SelectionMode="Multiple" Width="100%" Font-Names="Verdana" 
                    Font-Size="XX-Small"></asp:ListBox>
            </td>
            <td align="left" valign="top" style="font-family: Verdana; font-size: xx-small; margin-left: 10px">
                <b>
                <br />
                HOW TO SELECT MULTIPLE COST CENTRE CODES<br />
                </b>
                <br />
                To select two or more Cost Centre codes that are <b>not adjacent</b> in the 
                list, press and hold the CTRL key, then click on each Cost Centre code.<br />
                <br />
                To select a <b>block</b> of Cost Centres codes, click on the first code in the 
                block, press and hold the SHIFT key, then click on the last code in the block.</td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;</td>
            <td>
                &nbsp;</td>
            <td>
                <asp:Button ID="btnUpdate" runat="server" onclick="btnUpdate_Click" 
                    Text="Update" Width="100px" Enabled="False" />
            </td>
            <td align="right">
                &nbsp;</td>
            <td>
                &nbsp;</td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                &nbsp;<asp:Label ID="lblLegendUser0" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small">Cost Centres by User:</asp:Label>
                &nbsp;<asp:GridView ID="gvCostCentresByUser" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small" Width="100%" OnRowDataBound="gvCostCentresByUser_RowDataBound" 
                    AutoGenerateColumns="False" CellPadding="2">
                    <AlternatingRowStyle BackColor="#FFFFCC" />
                    <Columns>
                        <asp:TemplateField HeaderText="User">
                            <ItemTemplate>
                                <asp:HiddenField ID="hidUserKey" Value='<%# Bind("UserKey") %>' runat="server" />
                                <asp:Label ID="lblUserName" runat="server" Text='<%# Bind("UserName") %>'/>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Cost Centres">
                            <ItemTemplate>
                                <asp:Label ID="lblCostCentres" runat="server"/>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <RowStyle BackColor="#CCFFFF" />
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Label ID="lblLegendUsersByCostCentre" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small">Users by Cost Centre:</asp:Label>
                </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                <asp:GridView ID="gvUsersByCostCentre" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small" Width="100%" 
                    AutoGenerateColumns="False" CellPadding="2" OnRowDataBound="gvUsersByCostCentre_RowDataBound">
                    <AlternatingRowStyle BackColor="#FFFFCC" />
                    <Columns>
                        <asp:TemplateField HeaderText="Cost Centre" SortExpression="CostCentre">
                            <ItemTemplate>
                                <asp:Label ID="lblCostCentre" runat="server" Text='<%# Bind("CostCentre") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Users" SortExpression="User">
                            <ItemTemplate>
                                <asp:Label ID="lblUsers" runat="server"/>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EmptyDataTemplate>
                        *****
                        <asp:Label ID="lblLegendNothingConfigured" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small">nothing configured</asp:Label>
                        &nbsp;*****
                    </EmptyDataTemplate>
                    <RowStyle BackColor="#CCFFFF" />
                </asp:GridView>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlSpare" runat="server" Visible="false" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 2%">
                    &nbsp;
                </td>
                <td style="width: 26%">
                </td>
                <td style="width: 40%">
                    &nbsp;
                </td>
                <td style="width: 30%">
                    &nbsp;
                </td>
                <td style="width: 2%">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>

