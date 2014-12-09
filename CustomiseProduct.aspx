<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call RetrieveAccessParameters()
        End If
    End Sub
    
    Protected Sub RetrieveAccessParameters()
        ' <asp:HiddenField ID="username" runat="server" Value="geoffgilbert" />
        ' <asp:HiddenField ID="password" runat="server" Value="boxaxesu" />

        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT * FROM OnDemandCustomerAccess WHERE CustomerKey = " & Session("CustomerKey")
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                username.Value = oDataReader("username")
                password.Value = oDataReader("password")
                customer.Value = oDataReader("customer")
            Else
                WebMsgBox.Show("No access parameters found for product customisation. Please contact your Account Handler.")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrieveAccessParameters: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Customise Product</title>
</head>
<body onload="javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions(&quot;Button1&quot;, &quot;&quot;, false, &quot;&quot;, &quot;http://62.73.174.191/p4m_shop2/Workflow.do?start_workflow=Login&quot;, false, true))">
    <form id="form1" runat="server">
        <asp:HiddenField ID="username" runat="server" />
        <asp:HiddenField ID="password" runat="server" />
        <asp:HiddenField ID="customer" runat="server" />
        <asp:HiddenField ID="_action_name" runat="server" Value="Login" />
        <br />
        <br />
        <table style="width: 100%">
            <tr>
                <td align="center">
                    <asp:Button ID="Button1" runat="server" Text="Please wait, connecting you to the product customisation web site..."  PostBackUrl="http://62.73.174.191/p4m_shop2/Workflow.do?start_workflow=Login" />
                </td>
            </tr>
        </table>
        <br />
        &nbsp;
    </form>
</body>
</html>
