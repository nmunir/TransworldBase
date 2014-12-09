<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c As HttpCookie
        If (Request.Cookies("SprintLogon") Is Nothing) Then
            c = New HttpCookie("SprintLogon")
        Else
            c = Request.Cookies("SprintLogon")
        End If
        c.Values.Add("UserID", String.Empty)
        c.Values.Add("Password", String.Empty)
        c.Expires = DateTime.Now.AddYears(-30)
        Response.Cookies.Add(c)
        Label1.Text = "Auto Login Cookie Removed!"
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Large"></asp:Label></div>
    </form>
</body>
</html>
