<%@ Page Language="VB" %>
<%@ import Namespace="System.Drawing.Imaging" %>
<script runat="server">

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Copyright Jonathan Hare June 2004
    '   Product Manager: part of the web interface to e-log
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '
    '   MASTER COPY
    '
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~ Page Load ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Sub Page_Load(Source As Object, E As EventArgs)
        If Not IsPostBack Then

            Dim sImageFile As String = Request.QueryString("Image")

            Dim sVirtualJPGFolder As String = ConfigurationSettings.AppSettings("Virtual_JPG_URL")

            Image1.ImageUrl =  sVirtualJPGFolder & sImageFile

        End If
    End Sub

</script>
<html>
<head>
</head>
<body>
    <form runat="server">
        <asp:Image id="Image1" runat="server"></asp:Image>
        <!-- Insert content here -->
    </form>
</body>
</html>
