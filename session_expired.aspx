<%@ Page %>
<html>
<head>
    <title>Auto Log Out</title>
</head>
<body>
    <p align="center">
    </p>
    <p align="center">
    </p>
    <p align="center">
        <asp:Table id="Table1" runat="server" Width="100%">
            <asp:TableRow HorizontalAlign="Center">
                <asp:TableCell HorizontalAlign="Center">
                    <asp:Image runat="server" ID="Image1" ImageUrl="./images/icon_shutdown.gif"></asp:Image>
                    &nbsp; 
                    <asp:Label runat="server" ForeColor="Blue" ID="Label1" Font-Size="Small" Font-Names="Arial">As a security measure, you have been logged out of the system.</asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow HorizontalAlign="Center">
                <asp:TableCell HorizontalAlign="Center">
                    <asp:Label runat="server" ID="Label2" Font-Size="X-Small" Font-Names="Arial">This is normally due to inactivity.</asp:Label>
                    <asp:Label runat="server" ID="Label3" Font-Size="X-Small" Font-Names="Arial"> Please </asp:Label>
                    <asp:HyperLink runat="server" Font-Names="Arial" Font-Size="X-Small" NavigateUrl="default.aspx" ID="HyperLink1">click here</asp:HyperLink>
                    <asp:Label runat="server" ID="Label4" Font-Size="X-Small" Font-Names="Arial"> to logon again.</asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </p>
    <p align="left">
    </p>
</body>
</html>
