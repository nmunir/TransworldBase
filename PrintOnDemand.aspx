<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Print On Demand"
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
    <body>
        <form id="Form1" runat="Server">
        <main:Header id="ctlHeader" runat="server"></main:Header>
            <iframe frameborder = "0" 
                    runat="server"
                    width="100%"
                    height="600px"
                    id="PrintOnDemandIFrame" src="http://www.attem.co.uk" >Your browser does not support IFRAMEs</iframe>
    </form>
</body>
</html>

