<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<script runat="server">

    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Const CUSTOMER_LOVELLS As Integer = 663
    Const CUSTOMER_WURS As Int32 = 579
    Const CUSTOMER_WUIRE As Int32 = 686

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            cbAlwaysShowReportInNewWindow.Checked = pbNewWindow
            If cbAlwaysShowReportInNewWindow.Checked Then
                lnkbtnExpandAll.Visible = False
                lnkbtnCollapseAll.Visible = False
            End If
            pbIsProductOwner = CBool(Session("UserType").ToString.ToLower.Contains("owner"))
            Select Case pbIsProductOwner
                Case True                                                   ' assumes no customer with custom reports uses Product Owners
                    xdsReports.DataFile = "Reports_Default_ProductOwner.xml"
                Case False
                    If IsHyster() Then
                        xdsReports.DataFile = "Reports_Custom_Hyster.xml"
                    ElseIf IsYale() Then
                        xdsReports.DataFile = "Reports_Custom_Yale.xml"
                    ElseIf IsPorterNovelli() Then
                        xdsReports.DataFile = "Reports_Custom_PorterNovelli.xml"
                    ElseIf IsNHS() Then
                        xdsReports.DataFile = "Reports_Custom_NHS.xml"
                    ElseIf IsKODDFIS() Then
                        xdsReports.DataFile = "Reports_Custom_KODDFIS.xml"
                        'ElseIf IsFEXCO() Then
                        '    xdsReports.DataFile = "Reports_Custom_FEXCO.xml"
                    ElseIf IsWU() Then
                        xdsReports.DataFile = "Reports_Custom_WU.xml"
                        'ElseIf IsWUIRE() Then
                        '    xdsReports.DataFile = "Reports_Custom_WUIRE.xml"
                    ElseIf IsLovells() Then
                        xdsReports.DataFile = "Reports_Custom_LOVELLS.xml"
                    Else
                        xdsReports.DataFile = "Reports_Default_SuperUser.xml"
                    End If
            End Select
            ' xdsReports.DataFile = "Reports.xml"
        End If
        Call SetTitle()
        If cbAlwaysShowReportInNewWindow.Checked And pbCheckedPopupBlocking Then
            lnkbtnExpandAll.Visible = True
            lnkbtnCollapseAll.Visible = True
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Reports"
    End Sub
    
    Protected Function IsLovells() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsLovells = IIf(gbSiteTypeDefined, gsSiteType = "lovells", nCustomerKey = CUSTOMER_LOVELLS)
    End Function
    
    Protected Function IsHyster() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHyster = IIf(gbSiteTypeDefined, gsSiteType = "hyster", nCustomerKey = 77)
    End Function
   
    Protected Function IsYale() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsYale = IIf(gbSiteTypeDefined, gsSiteType = "yale", nCustomerKey = 680)
    End Function
   
    Protected Function IsFEXCO() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsFEXCO = IIf(gbSiteTypeDefined, gsSiteType = "fexco", nCustomerKey = 579)
    End Function
   
    Protected Function IsWU() As Boolean
        Dim arrWU() As Integer = {CUSTOMER_WURS, CUSTOMER_WUIRE}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWU = IIf(gbSiteTypeDefined, gsSiteType = "wu", Array.IndexOf(arrWU, nCustomerKey) >= 0)
    End Function

    Protected Function IsWUIRE() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWUIRE = IIf(gbSiteTypeDefined, gsSiteType = "wuire", nCustomerKey = 686)
    End Function
   
    Protected Function IsKODDFIS() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsKODDFIS= IIf(gbSiteTypeDefined, gsSiteType = "koddfis", nCustomerKey = 541)
    End Function
   
    Protected Function IsPorterNovelli() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsPorterNovelli = IIf(gbSiteTypeDefined, gsSiteType = "porternovelli", nCustomerKey = 314)
    End Function
   
    Protected Function IsNHS() As Boolean
        Dim arrCustomerNHS() As Integer = {260, 484, 488, 485, 487, 483, 486, 265}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsNHS = IIf(gbSiteTypeDefined, gsSiteType = "nhs", Array.IndexOf(arrCustomerNHS, nCustomerKey) >= 0)
    End Function

    Protected Sub tvReports_SelectedNodeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tv As TreeView = sender
        lblReportTitle.Text = tv.SelectedNode.Text
        lblReportDescription.Text = tv.SelectedNode.Value
        
        hyplnkShowInNewWindow.NavigateUrl = tv.SelectedNode.Target
        If cbAlwaysShowReportInNewWindow.Checked AndAlso tv.SelectedNode.Target = String.Empty And Not pbCheckedPopupBlocking Then
            Call SetPopupsBlockedCheck()
            pbCheckedPopupBlocking = True
        End If
        If cbAlwaysShowReportInNewWindow.Checked AndAlso tv.SelectedNode.Target <> String.Empty Then
            Call SetPopup(tv.SelectedNode.Target)
        Else
            ReportIFrame.Attributes("src") = tv.SelectedNode.Target
        End If
        If tv.SelectedNode.Depth = 2 Then
            hyplnkShowInNewWindow.Enabled = True
        Else
            hyplnkShowInNewWindow.Enabled = False
        End If
    End Sub

    Protected Sub SetPopupsBlockedCheck()
        Dim sCSName As String = "PopupsBlockedCheck"
        Dim tpCSType As Type = Me.GetType()
        Dim csmClientScriptManager As ClientScriptManager = Page.ClientScript
        If (Not csmClientScriptManager.IsStartupScriptRegistered(tpCSType, sCSName)) Then
            csmClientScriptManager.RegisterStartupScript(tpCSType, sCSName, "IsPopupBlocked();", True)
        End If
    End Sub

    Protected Sub SetPopup(ByVal sTarget As String)
        Dim sCSname As String = "PopupScript"
        Dim cstype As Type = Me.GetType()
        Dim csmClientScriptManager As ClientScriptManager = Page.ClientScript
        If (Not csmClientScriptManager.IsStartupScriptRegistered(cstype, sCSname)) Then
            Dim sCStext As String = "window.open('" & sTarget & "','Report','top=10,left=10,width=1000,height=600,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes');"
            csmClientScriptManager.RegisterStartupScript(cstype, sCSname, sCStext, True)
        End If
    End Sub
    
    Protected Sub CreateSprintConfigCookie()
        Dim c As HttpCookie = New HttpCookie("SprintConfig")
        c.Values.Add("RPT_NewWindow", "False")
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub
    
    Protected Sub UpdateSprintConfigCookieNewWindow(ByVal sNewWindow As String)
        Dim c As HttpCookie = New HttpCookie("SprintConfig")
        c.Values.Add("RPT_NewWindow", sNewWindow)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub

    Property pbNewWindow() As Boolean
        Get
            Dim oViewState As Object = ViewState("RPT_NewWindow")
            Dim bNewWindow As Boolean = False
            If oViewState Is Nothing Then
                If Request.Cookies("SprintConfig") Is Nothing Then
                    Call CreateSprintConfigCookie()
                    ViewState("RPT_NewWindow") = False
                Else
                    If Request.Cookies("SprintConfig")("RPT_NewWindow") Is Nothing Then
                        Call UpdateSprintConfigCookieNewWindow("False")
                        ViewState("RPT_NewWindow") = False
                    Else
                        ViewState("RPT_NewWindow") = CBool(Request.Cookies("SprintConfig")("RPT_NewWindow"))
                        bNewWindow = ViewState("RPT_NewWindow")
                    End If
                End If
            Else
                bNewWindow = CBool(oViewState)
            End If
            Return bNewWindow
        End Get
        Set(ByVal Value As Boolean)
            Call UpdateSprintConfigCookieNewWindow(CStr(Value))
            ViewState("RPT_NewWindow") = Value
        End Set
    End Property
    
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("RPT_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("RPT_IsProductOwner") = Value
        End Set
    End Property
   
    Property pbCheckedPopupBlocking() As Boolean
        Get
            Dim o As Object = ViewState("RPT_CheckedPopupBlocking")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("RPT_CheckedPopupBlocking") = Value
        End Set
    End Property
   
    Protected Sub cbAlwaysShowReportInNewWindow_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        pbNewWindow = cb.Checked
        ReportIFrame.Attributes("src") = String.Empty
        lblReportTitle.Text = String.Empty
        lblReportDescription.Text = String.Empty
        If cb.Checked Then
            tvReports.CollapseAll()
            pbCheckedPopupBlocking = False
            lnkbtnExpandAll.Visible = False
            lnkbtnCollapseAll.Visible = False
        Else
            lnkbtnExpandAll.Visible = True
            lnkbtnCollapseAll.Visible = True
        End If
    End Sub
    
    Protected Sub lnkbtnExpandAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tvReports.ExpandAll()
    End Sub
    
    Protected Sub lnkbtnCollapseAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tvReports.CollapseAll()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" runat="Server">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table style="width: 100%" cellpadding="0" cellspacing="0">
            <tr class="bar_reports">
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap" align="right">
                </td>
            </tr>
        </table>
        &nbsp;
        <asp:XmlDataSource ID="xdsReports" runat="server" />
        <table width="100%" cellpadding="10">
            <tr>
                <td style="width: 250px; height: 165px;" valign="top">
                    <asp:TreeView ID="tvReports" DataSourceID="xdsReports" Target="test" runat="server"
                        OnSelectedNodeChanged="tvReports_SelectedNodeChanged" ExpandDepth="1" Style="color: blue"
                        CssClass="TreeView">
                        <DataBindings>
                            <asp:TreeNodeBinding DataMember="Reports" Depth="0" TextField="Text" ToolTipField="ToolTip"
                                SelectAction="SelectExpand" ValueField="Value" />
                            <asp:TreeNodeBinding DataMember="ReportClass" Depth="1" TextField="Text" ToolTipField="ToolTip"
                                SelectAction="SelectExpand" ValueField="Value" />
                            <asp:TreeNodeBinding DataMember="Report" Depth="2" TextField="Text" ValueField="Value"
                                TargetField="URL" />
                        </DataBindings>
                        <HoverNodeStyle ForeColor="Black" />
                        <NodeStyle Font-Names="Verdana,Sans-Serif" Font-Size="XX-Small" />
                    </asp:TreeView>
                    <br />
                    <asp:LinkButton ID="lnkbtnExpandAll" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Silver" OnClick="lnkbtnExpandAll_Click">expand all</asp:LinkButton>
                    <asp:LinkButton ID="lnkbtnCollapseAll" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Silver" OnClick="lnkbtnCollapseAll_Click">collapse all</asp:LinkButton></td>
                <td valign="top" style="width: 633px; height: 165px">
                    <asp:Panel ID="pnlDescription" runat="server" BorderStyle="None" BorderWidth="1px" Height="100%"
                        Width="100%" BorderColor="#E0E0E0">
                        <p align="center">
                            <asp:Label ID="lblReportTitle" runat="server" Font-Bold="True" Font-Size="Small" Font-Names="Verdana"></asp:Label>
                        </p>
                        <asp:Label ID="lblReportDescription" runat="server" Font-Size="XX-Small" Font-Names="Verdana"></asp:Label>
                     </asp:Panel>
                </td>
                <td valign="top" style="height: 165px">
                </td>
            </tr>
            <tr>
                <td colspan="3" align="right">
                    <asp:HyperLink ID="hyplnkShowInNewWindow" runat="server" Target="_blank"
                        Font-Size="XX-Small" Font-Names="Verdana">Show report in a new window</asp:HyperLink>
                    &nbsp;&nbsp;<asp:CheckBox ID="cbAlwaysShowReportInNewWindow" runat="server" AutoPostBack="True"
                        Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbAlwaysShowReportInNewWindow_CheckedChanged"
                        Text="(always)" />&nbsp;</td>
            </tr>
            <tr>
                <td colspan="3" valign="top" style="height: 20000px">
                    <iframe frameborder="0" runat="server" width="100%" height="100000px" id="ReportIFrame"
                        src="">Your browser does not support IFRAMEs</iframe>
                </td>
            </tr>
        </table>
        <br />

        <script type="text/javascript">
           function IsPopupBlocked() {
             var strNewURL = "dummy.htm";
             var Strfeature = "top=10,left=10,width=0,height=0,status=no,toolbar=no,address=no,menubar=no,resizable=no,scrollbars=no" ;
             var WindowOpen = window.open (strNewURL,"MainWindow",Strfeature);
             try {
               var obj = WindowOpen.name;
               WindowOpen.close();
             }
             catch(e){
               alert("Grrr...your browser is blocking pop-up windows!\n\nTo show reports in a new window, please disable your pop-up window blocker for this site and try again.\n\nTypically, pop-up windows are blocked by the browser itself (in IE7 go to Tools | Pop-up Blocker) or by a 3rd party toolbar such as Google toolbar or Yahoo! toolbar.\n\nContact your system administrator for further advice. ");
             }
           }
        </script>
    </form>
</body>
</html>

