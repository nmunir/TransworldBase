<%@ Page %>
<html>
<head>
    <title>Error</title>
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
                    <asp:Label runat="server" ForeColor="Blue" ID="Label1" Font-Size="Small" Font-Names="Arial">Our apologies, an error has occurred on the page you are requesting.</asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            <asp:TableRow HorizontalAlign="Center">
                <asp:TableCell HorizontalAlign="Center">
                    <asp:Label runat="server" ID="Label2" Font-Size="X-Small" Font-Names="Arial">Please contact customer services for assistance.</asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </p>
    <p align="left">
    </p>
</body>
</html>
