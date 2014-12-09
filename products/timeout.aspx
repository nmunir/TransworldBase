<%@ Page %>
<html>
<head>
    <title>Security Timeout</title>
</head>
<body>
    <p>
    </p>
    <p>
    </p>
        <table style="width:100%">
            <tr align="center">
                <td align="center">
                    <asp:Image runat="server" ID="Image1" ImageUrl="./images/icon_shutdown.gif" />
                    &nbsp; 
                    <asp:Label runat="server" ForeColor="Blue" ID="Label1" Font-Size="Small" Font-Names="Verdana">Your session has been terminated because of inactivity. </asp:Label><asp:HyperLink
                        ID="HyperLink1" runat="server" NavigateUrl="~/default.aspx" Font-Size="Small" Font-Names="Verdana" ForeColor="Blue">Restart session.</asp:HyperLink>
                </td>
            </tr>
            <tr align="center">
                <td>
                </td>
            </tr>
        </table>
</body>
</html>
