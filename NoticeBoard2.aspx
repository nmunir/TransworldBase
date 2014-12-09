<%@ Page Language="VB" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.XML" %>
<%@ Import Namespace="System.ServiceModel.Syndication" %>
<%@ Import Namespace="System.Net" %>

<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
  
    Protected Sub Page_Load()
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
            
        If Not IsPostBack Then
            Call LoadControls()
        End If
        Call SetTitle()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Notice Board"
    End Sub
    
    Private Sub LoadControls()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetPagePanelControlsFromCustomerKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        Try
            ltrLeftHTML.Text = String.Empty
            ltrRightHTML.Text = String.Empty
            ltrCentreHTML.Text = String.Empty
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count = 1 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                ltrLeftHTML.Text = dr("LeftHTML").ToString()
                ltrRightHTML.Text = dr("RightHTML").ToString()
                ltrCentreHTML.Text = dr("CentreHTML").ToString()
                
                If IsTokenExist(ltrLeftHTML) Then
                    FindControlByTags(ltrLeftHTML)
                End If
                If IsTokenExist(ltrRightHTML) Then
                    FindControlByTags(ltrRightHTML)
                End If
                If IsTokenExist(ltrCentreHTML) Then
                    FindControlByTags(ltrCentreHTML)
                End If
            Else
                If oDataTable.Rows.Count > 1 Then
                    WebMsgBox.Show("LoadControls: more than one item found - please inform development")
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("LoadControls: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function IsTokenExist(ByVal lbl As Label) As Boolean
        
        IsTokenExist = True
        If Not lbl.Text.ToLower.Contains("[rotator:") And Not lbl.Text.ToLower.Contains("[textbox:") Then
            IsTokenExist = False
        End If
        
    End Function
    
    Protected Function FindControlByTags(ByVal lblHTML As Label) As String
        'Dim regularExpressionPattern As String = "\[(.*?)\]"
        Dim regularExpressionPattern As String = "\[[a-zA-Z0-9]*.*\]"
        Dim re As New Regex(regularExpressionPattern)
        For Each oMatch In re.Matches(lblHTML.Text)
            CreateDynamicControl(oMatch.ToString(), lblHTML)
        Next
        Return lblHTML.Text
    End Function
    
    Protected Function RemoveHTMLTags(ByVal strHTML As String) As String
        Return Regex.Replace(strHTML, "<(.|\n)*?>", String.Empty)
    End Function
    
    Protected Sub CreateDynamicControl(ByVal controlTags As String, ByVal ctllbl As Label)
        Dim sDataSourceTag As String = String.Empty
        If controlTags.ToLower.Contains("rotator") Then
            Dim strTemplateControls As String() = controlTags.Split(":")
            Dim rr As New RadRotator
            If strTemplateControls.Length > 0 Then
                rr.ItemTemplate = New RadRotatorTemplate(strTemplateControls)
            Else
                rr.ItemTemplate = New RadRotatorTemplate
            End If
            

            If controlTags.ToLower.Contains(";") Then
                Dim sRotatorProperties As String() = controlTags.Split(";")
                For Each config As String In sRotatorProperties
                    If config.ToLower().Contains("height=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.Height = Convert.ToInt32(sValue)
                        Else
                            rr.Height = 100
                        End If
                    ElseIf config.ToLower().Contains("width=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.Width = Convert.ToInt32(sValue)
                        Else
                            rr.Width = 400
                        End If
                    ElseIf config.ToLower().Contains("scrolldirection=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "up" Then
                                rr.ScrollDirection = RotatorScrollDirection.Up
                            ElseIf sValue.ToLower() = "down" Then
                                rr.ScrollDirection = RotatorScrollDirection.Down
                            ElseIf sValue.ToLower() = "left" Then
                                rr.ScrollDirection = RotatorScrollDirection.Left
                            ElseIf sValue.ToLower() = "right" Then
                                rr.ScrollDirection = RotatorScrollDirection.Right
                            Else
                                rr.ScrollDirection = RotatorScrollDirection.Up
                            End If
                        End If
                    ElseIf config.ToLower().Contains("scrollduration=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.ScrollDuration = Convert.ToInt32(sValue)
                        Else
                            rr.ScrollDuration = 3000
                        End If
                    ElseIf config.ToLower().Contains("data source") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            sDataSourceTag = sValue
                        End If
                    ElseIf config.ToLower.Contains("rotatortype=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "buttons" Then
                                rr.RotatorType = RotatorType.Buttons
                            ElseIf sValue.ToLower() = "automaticadvance" Then
                                rr.RotatorType = RotatorType.AutomaticAdvance
                            ElseIf sValue.ToLower() = "buttonsover" Then
                                rr.RotatorType = RotatorType.ButtonsOver
                            ElseIf sValue.ToLower() = "carousel" Then
                                rr.RotatorType = RotatorType.Carousel
                            ElseIf sValue.ToLower() = "carouselbuttons" Then
                                rr.RotatorType = RotatorType.CarouselButtons
                            ElseIf sValue.ToLower() = "coverflow" Then
                                rr.RotatorType = RotatorType.CoverFlow
                            ElseIf sValue.ToLower() = "coverflowbuttons" Then
                                rr.RotatorType = RotatorType.CoverFlowButtons
                            ElseIf sValue.ToLower() = "slideshow" Then
                                rr.RotatorType = RotatorType.SlideShow
                            ElseIf sValue.ToLower() = "slideshowbuttons" Then
                                rr.RotatorType = RotatorType.SlideShowButtons
                            Else
                                rr.RotatorType = RotatorType.Carousel
                            End If
                        End If
                    ElseIf config.ToLower.Contains("animationtype=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "none" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.None
                            ElseIf sValue.ToLower() = "fade" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Fade
                            ElseIf sValue.ToLower() = "pulse" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Pulse
                            ElseIf sValue.ToLower() = "crossfade" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.CrossFade
                            End If
                        End If
                    ElseIf config.ToLower.Contains("css=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            rr.CssClass = sValue
                        End If
                    End If
                Next
            End If
            
           

            'rr.Attributes.Add(  
            'rr.CssClass = "rotator"
            rr.ItemHeight = 50
            rr.DataSource = BindRotator(sDataSourceTag)
            rr.DataBind()
            ctllbl.Controls.Add(rr)
        End If
        
        If controlTags.ToLower.Contains("textbox:") Then
            Dim stringSeparators() As String = {"[textbox:"}
            Dim strTemplateControls As String() = controlTags.Split(stringSeparators, StringSplitOptions.None)
            If strTemplateControls.Length > 1 Then
                Dim lbl As New Label
                lbl.Text = strTemplateControls(1).ToString.Replace("]", "")
                ctllbl.Controls.Add(lbl)
            End If
        End If
        
        
        
    End Sub
    
    Protected Function BindRotator(ByVal sDataSourceTag As String) As DataTable
        
        Dim oDataTable As New DataTable
        BindRotator = Nothing
        If sDataSourceTag.ToLower = "rss" Then
            BindRotator = ExtractFromRSSFeed()
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetContentFromCustomerKey", oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DataSourceTag", SqlDbType.VarChar))
            oAdapter.SelectCommand.Parameters("@DataSourceTag").Value = sDataSourceTag
            Try
                oAdapter.Fill(oDataTable)
                BindRotator = oDataTable
            Catch ex As Exception
                WebMsgBox.Show("GetPageContent: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        End If
    End Function
    
    Protected Sub RemoveAutoLogonCookie()
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
    End Sub
    
#Region "Rss Feed"
    
    Protected Function CreateProductsDataTable() As DataTable
        Dim oDataTable As New DataTable("RotatorRssFeed")
        oDataTable.Columns.Add(New DataColumn("Title", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Content", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("ImageTag", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Date", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("BaseUri", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Categories", GetType(String)))
        CreateProductsDataTable = oDataTable
    End Function
    
    Protected Function ExtractFromRSSFeed() As DataTable
        
        Dim dt As DataTable = CreateProductsDataTable()
        Dim dr As DataRow
        Dim nRssCount As Integer = 0
        Dim sRssUrl As String = String.Empty
        Dim nCustomerKey As Integer = Convert.ToInt64(Session("CustomerKey"))
        If nCustomerKey > 0 Then
            Dim sQuery As String = "select RssUrl, RssCount from RotatorRssFeed where CustomerKey = " & nCustomerKey
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
            If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
                dr = oDataTable.Rows(0)
                If Not IsDBNull(dr("RssCount")) Then
                    nRssCount = Convert.ToInt32(dr("RssCount"))
                End If
                If Not IsDBNull(dr("RssUrl")) Then
                    sRssUrl = dr("RssUrl").ToString()
                End If
            End If
        End If
        
        
        Dim sf As SyndicationFeed = LoadRSSFeed(sRssUrl)
        If Not sf Is Nothing Then
            For Each item As SyndicationItem In sf.Items
                'Dim item As SyndicationItem = sf.Items(i)
                dr = dt.NewRow
                dr("Title") = item.Title.Text
                dr("Date") = Convert.ToDateTime(item.PublishDate.LocalDateTime).ToString("dd-MMM-yyyy HH:mm")
                dr("Content") = item.Summary.Text
                If Not item.Links Is Nothing Then
                    For Each sl As SyndicationLink In item.Links
                        dr("BaseURI") = sl.Uri.AbsoluteUri
                    Next
                End If
                Dim bFirst As Boolean = True
                For Each sc As SyndicationCategory In item.Categories
                    If Not bFirst Then
                        dr("Categories") += ", "
                    End If
                    'dr("Categories") += sc.Name
                    dr("Categories") += sc.Name
                    bFirst = False
                Next
                dt.Rows.Add(dr)
                'If i = CInt(nRssCount) - 1 Then
                '    Exit For
                'End If
            Next
        End If
        ExtractFromRSSFeed = dt
    End Function
    
    Protected Function LoadRSSFeed(ByVal sRssUrl As String) As SyndicationFeed
        Try
            Using reader As XmlReader = XmlReader.Create(sRssUrl)
                LoadRSSFeed = SyndicationFeed.Load(reader)
            End Using
        Catch ex As WebException
        Catch ex As XmlException
        Catch ex As Exception
        End Try
    End Function
    
#End Region
    
    
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <link rel="Stylesheet" type="text/css" href="~/css/sprint_rotator.css" /> 
    <title>Notice Board</title>
    <style type="text/css">
        .image
        {
            vertical-align: middle;
        }                
       
    </style>    
</head>
<body>
    <form id="frmNoticeBoard" runat="Server">    
    <main:Header ID="ctlHeader" runat="server">
    </main:Header>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_noticeboard">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <table style="width: 100%" runat="server" id="tbl_main">
        <tr>
            <td>
                <table id="tblHeader" style="width: 100%">
                    <tr>
                        <td id="tdHeader" width="100%" runat="server">
                            &nbsp;
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table style="width: 100%" id="tblPageSections">
        <tr>
            <td valign="top" width="30%" runat="server" id="left">                
                <asp:Label ID="ltrLeftHTML" runat="server"></asp:Label>
            </td>
            <td valign="top" width="30%" runat="server" id="centre">
                <asp:Label ID="ltrCentreHTML" runat="server"></asp:Label>
            </td>
            <td valign="top" width="40%" runat="server" id="rtr">
                <asp:Label ID="ltrRightHTML" runat="server"></asp:Label>
            </td>
    </table>
    <div id="div_bottom" runat="server">
    </div>
    </form>
</body>
</html>