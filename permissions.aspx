<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' 	EXEC spASPNET_Email_AddToQueue 'AUTHORISATION RESP', @CustomerKey, NULL, NULL, NULL, @UserEmailAddr, @MsgSubject, @MsgBody, @HTMLMsgBody, 0

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Const GVCOL_CONTROLS As Integer = 0
    Const GVCOL_PERMISSIONS As Integer = 1
    Const GVCOL_NOTINTERNALUSER As Integer = 2
    Const GVCOL_USERID As Integer = 3
    Const GVCOL_NAME As Integer = 4
    Const GVCOL_CUSTOMERNAME As Integer = 5
    
    Dim COLOUR_ODD As System.Drawing.Color = System.Drawing.Color.White
    Dim COLOUR_EVEN As System.Drawing.Color = System.Drawing.Color.AntiqueWhite
    Dim gsLastCustomer As String = String.Empty
    Dim gnCurrentColour As System.Drawing.Color = COLOUR_EVEN

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Page.MaintainScrollPositionOnPostBack = True
        ' tbCurrentPassword.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitPasswordChangeRequest.ClientID + "')")
    End Sub
    
    Protected Sub HideAllPanels()
    End Sub
    
    Protected Sub btnRetryTransmissionNow_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String = "UPDATE EmailQueue SET NextRetryOn = '8-jun-2008' WHERE NextRetryOn IS NOT NULL"
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            gvPermissions.DataSource = oDataTable
            gvPermissions.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("RetryTransmissionNow: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnRetrievePermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RetrievePermissions()
        gvPermissions.Visible = True
    End Sub
    
    Protected Sub RetrievePermissions()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT UserPermissions 'Permissions', up.Customer 'Not Internal User', UserId 'User', FirstName + ' ' + LastName 'Name', CustomerAccountCode 'Customer' FROM UserProfile up INNER JOIN Customer c ON up.CustomerKey = c.CustomerKey WHERE ((UserPermissions > 0) OR (Customer = 0)) AND Status = 'Active' ORDER BY c.CustomerKey"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            gvPermissions.DataSource = oDataTable
            gvPermissions.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrievePermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub gvPermissions_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRowEventArgs = e
        Dim row As GridViewRow = e.Row
        row.BackColor = Drawing.Color.AntiqueWhite
        If gsLastCustomer = row.Cells(GVCOL_CUSTOMERNAME).Text Then
            row.BackColor = gnCurrentColour
        Else
            Call FlipColour()
            row.BackColor = gnCurrentColour
            gsLastCustomer = row.Cells(GVCOL_CUSTOMERNAME).Text
        End If
    End Sub
    
    Protected Sub FlipColour()
        If gnCurrentColour = COLOUR_ODD Then
            gnCurrentColour = COLOUR_EVEN
        Else
            gnCurrentColour = COLOUR_ODD
        End If
    End Sub
    
    Protected Sub btnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim gvr As GridViewRow = b.NamingContainer
        Dim tb As TextBox = gvr.FindControl("tbNewValue")
        If IsNumeric(tb.Text) Then
            Call Update(b.CommandArgument, tb.Text)
        Else
            WebMsgBox.Show("Number required!")
        End If
    End Sub
    
    Protected Sub Update(ByVal sUserId As String, ByVal nNewValue As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "UPDATE UserProfile SET UserPermissions = " & nNewValue & " WHERE UserId = '" & sUserId.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in Update: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSetInternalUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Call SetInternalUser(b.CommandArgument)
    End Sub
    
    Protected Sub SetInternalUser(ByVal sUserId As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "UPDATE UserProfile SET Customer = 0 WHERE UserId = '" & sUserId.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetInternalUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnClearInternalUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Call ClearInternalUser(b.CommandArgument)
    End Sub
    
    Protected Sub ClearInternalUser(ByVal sUserId As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "UPDATE UserProfile SET Customer = 1 WHERE UserId = '" & sUserId.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ClearInternalUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnHidePermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvPermissions.Visible = False
    End Sub
    
    Protected Sub btnShowCustomersWithNoAccountHandler_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCustomersWithNoAccountHandler()
        gvCustomers.Visible = True
    End Sub
    
    Protected Sub ShowCustomersWithNoAccountHandler()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT CustomerAccountCode FROM Customer WHERE CustomerKey NOT IN (SELECT CustomerKey FROM UserProfile WHERE (UserPermissions & 1 > 0) AND Status = 'Active') AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        Try
            oAdapter.Fill(oDataTable)
            gvCustomers.DataSource = oDataTable
            gvCustomers.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrievePermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnHideCustomers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvCustomers.Visible = False
    End Sub
    
</script>
<html xmlns=" http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Permissions Utility</title>
    <link href="sprint.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmUserIdApplication" runat="server">
        <strong>&nbsp;User Permissions Utility<br />
        </strong>
        <br />
        <asp:Panel id="pnlMain" runat="server" visible="True" Width="100%">
            <br />
            &nbsp;<asp:Button ID="btnRetrievePermissions" runat="server" Text="show permissions" OnClick="btnRetrievePermissions_Click" />
            &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; 
            <asp:LinkButton ID="lnkbtnHidePermissions" runat="server" OnClick="lnkbtnHidePermissions_Click">hide permissions</asp:LinkButton>
            &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            <asp:Button ID="btnShowCustomersWithNoAccountHandler" runat="server" OnClick="btnShowCustomersWithNoAccountHandler_Click"
                Text="show customers with no account handler" />
            &nbsp; &nbsp; &nbsp;&nbsp;
            <asp:LinkButton ID="lnkbtnHideCustomers" runat="server" OnClick="lnkbtnHideCustomers_Click">hide customers</asp:LinkButton><br />
            <br />
            <asp:GridView ID="gvPermissions" runat="server" CellPadding="2" Font-Names="Verdana"
                Font-Size="XX-Small" Width="100%" OnRowDataBound="gvPermissions_RowDataBound" >
                <Columns>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:TextBox ID="tbNewValue" runat="server" Width="50px"></asp:TextBox>
                            <asp:Button ID="btnUpdate" runat="server" OnClick="btnUpdate_Click" CommandArgument='<%# DataBinder.Eval(Container, "DataItem.User") %>' Text="update" Width="50px" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnSetInternalUser" runat="server" Text="set internal user" Width="110px" OnClick="btnSetInternalUser_Click" CommandArgument='<%# DataBinder.Eval(Container, "DataItem.User") %>'  />
                            <asp:Button ID="btnClearInternalUser" runat="server" OnClick="btnClearInternalUser_Click"
                                Text="clear internal user" Width="110px" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
            &nbsp;&nbsp;&nbsp;&nbsp;<asp:GridView ID="gvCustomers" runat="server" CellPadding="2" Font-Names="Verdana"
                Font-Size="XX-Small" Width="100%" >
                <Columns>
                </Columns>
            </asp:GridView>
            <br />
            </asp:Panel>
        </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>

</body>
</html>