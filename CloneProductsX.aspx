<%@ Page Language="VB" validaterequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.XML" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Collections.Generic" %>

<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gsXMLRotatorConfigFilePath As String
    Dim gsXMLNewsContentFilePath As String
    
    Protected Sub Page_Load()
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        gsXMLRotatorConfigFilePath = ".\rotator\news_config" & Session("SiteKey") & ".xml"
        gsXMLNewsContentFilePath = ".\rotator\news" & Session("SiteKey") & ".xml"
            
        If Not IsPostBack Then
            Call InitRotator()
            oRotator.XmlContentFile = gsXMLNewsContentFilePath
            LHSTextBlock.DataBind()
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
    
    Protected Sub InitRotator()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "NB1Content"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        If oDataTable.Rows.Count = 0 Then
            Call RemoveAutoLogonCookie()
            Server.Transfer("session_expired.aspx")
        End If

        Dim dr As DataRow = oDataTable.Rows(0)
        Dim sTemp As String
        
        oRotator.Visible = dr("NB1RtrVisible")
        oRotator.Loop = dr("NB1RtrContinuousLoop")
        oRotator.PauseOnMouseOver = dr("NB1RtrPauseOnMouseOver")
        sTemp = dr("NB1RtrScrollDirection")
        If sTemp = "Up" Then
            oRotator.ScrollDirection = ComponentArt.Web.UI.ScrollDirection.Up
        ElseIf sTemp = "Left" Then
            oRotator.ScrollDirection = ComponentArt.Web.UI.ScrollDirection.Left
        Else
            oRotator.ScrollDirection = ComponentArt.Web.UI.ScrollDirection.Up
        End If
        oRotator.SlidePause = dr("NB1RtrSlidePause")
        oRotator.ScrollInterval = dr("NB1RtrScrollInterval")
        sTemp = dr("NB1RtrRotationType")
        If sTemp = "ContentScroll" Then
            oRotator.RotationType = ComponentArt.Web.UI.RotationType.ContentScroll
        ElseIf sTemp = "SlideShow" Then
            oRotator.RotationType = ComponentArt.Web.UI.RotationType.SlideShow
        End If
        sTemp = dr("NB1RtrSmoothScrollSpeed")
        If sTemp = "Slow" Then
            oRotator.SmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Slow
        ElseIf sTemp = "Medium" Then
            oRotator.SmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Medium
        ElseIf sTemp = "Fast" Then
            oRotator.SmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Fast
        End If
        sTemp = dr("NB1RtrShowEffect")
        If sTemp = "None" Then
            oRotator.ShowEffect = ComponentArt.Web.UI.RotationEffect.None
        ElseIf sTemp = "Fade" Then
            oRotator.ShowEffect = ComponentArt.Web.UI.RotationEffect.Fade
        ElseIf sTemp = "Pixelate" Then
            oRotator.ShowEffect = ComponentArt.Web.UI.RotationEffect.Pixelate
        ElseIf sTemp = "Dissolve" Then
            oRotator.ShowEffect = ComponentArt.Web.UI.RotationEffect.Dissolve
        ElseIf sTemp = "GradientWipe" Then
            oRotator.ShowEffect = ComponentArt.Web.UI.RotationEffect.GradientWipe
        End If
        oRotator.ShowEffectDuration = dr("NB1RtrShowEffectDuration")
        sTemp = dr("NB1RtrHideEffect")
        If sTemp = "None" Then
            oRotator.HideEffect = ComponentArt.Web.UI.RotationEffect.None
        ElseIf sTemp = "Fade" Then
            oRotator.HideEffect = ComponentArt.Web.UI.RotationEffect.Fade
        ElseIf sTemp = "Pixelate" Then
            oRotator.HideEffect = ComponentArt.Web.UI.RotationEffect.Pixelate
        ElseIf sTemp = "Dissolve" Then
            oRotator.HideEffect = ComponentArt.Web.UI.RotationEffect.Dissolve
        ElseIf sTemp = "GradientWipe" Then
            oRotator.HideEffect = ComponentArt.Web.UI.RotationEffect.GradientWipe
        End If
        oRotator.HideEffectDuration = dr("NB1RtrHideEffectDuration")
        Call ParseAttributes(dr("NB1_AllAttr") & String.Empty, tbl_main)
        div_top.InnerHtml = dr("NB1_TopContent") & String.Empty
        Call ParseAttributes(dr("NB1_TopAttr") & String.Empty, div_top)
        LHSTextBlock.InnerHtml = dr("NB1_BodyContent")
        Call ParseAttributes(dr("NB1LeftAttr") & String.Empty, left)
        Call ParseAttributes(dr("NB1CentreAttr") & String.Empty, td_body)
        Call ParseAttributes(dr("NB1_BodyAttr") & String.Empty, LHSTextBlock)
        Call ParseAttributes(dr("NB1RightAttr") & String.Empty, right)
        div_bottom.InnerHtml = dr("NB1_BottomContent") & String.Empty
        Call ParseAttributes(dr("NB1_BottomAttr") & String.Empty, div_bottom)
        Call ParseAttributes(dr("NB1RtrAttr") & String.Empty, rtr)
    End Sub

    Protected Sub ParseAttributes(ByVal sAttributesField As String, ByVal hcDestinationField As HtmlControl)
        If sAttributesField <> String.Empty Then
            Dim sAttributes() As String = sAttributesField.Split(";")
            Dim dictAttributes As New Dictionary(Of String, String)
            For Each sAttributeKeyValue As String In sAttributes
                Dim sAttribute() As String = sAttributeKeyValue.Split(":")
                If sAttribute.GetUpperBound(0) = 1 Then
                    dictAttributes.Add(sAttribute(0), sAttribute(1))
                End If
            Next
            For Each kv As KeyValuePair(Of String, String) In dictAttributes
                hcDestinationField.Style(kv.Key) = kv.Value
            Next
        End If
    End Sub
    
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

    Property sAdRotatorImageFolder() As String
        Get
            Dim o As Object = ViewState("NB_AdRotatorImageFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("NB_AdRotatorImageFolder") = Value
        End Set
    End Property
    
    Property sVirtualAdrotatorThumbFolder() As String
        Get
            Dim o As Object = ViewState("NB_VirtualAdrotatorThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("NB_VirtualAdrotatorThumbFolder") = Value
        End Set
    End Property
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Notice Board</title>
</head>
<body>
    <form id="frmNoticeBoard" runat="Server">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table style="width: 100%" cellpadding="0" cellspacing="0">
            <tr class="bar_noticeboard">
                <td style="width: 50%; white-space:nowrap">
                </td>
                <td style="width: 50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <div id="div_top" style="width: 100%" runat="server"></div>
            <table style="width: 100%" runat="server" id="tbl_main">
                <tr>
                    <td runat="server" id="left"></td>
                    <td valign="top" runat="server" id="td_body">
                        <div id="LHSTextBlock" runat="server"></div>
                    </td>
                    <td runat="server" id="right"></td>
                    <td align="left" valign="top" runat="server" id="rtr">
                        <ComponentArt:Rotator ID="oRotator" runat="server" Width="100%" Height="350">
                            <SlideTemplate>
                                <span class="NewsDate">
                                    <br />
                                    <%# DataBinder.Eval(Container.DataItem, "Date") %>
                                </span><span class="NewsTitle">
                                    <br />
                                    <%# DataBinder.Eval(Container.DataItem, "Title") %>
                                </span><span class="NewsText">
                                    <br />
                                    <%# DataBinder.Eval(Container.DataItem, "Text") %>
                                </span>
                            </SlideTemplate>
                        </ComponentArt:Rotator>
                    </td>
                </tr>
            </table>
            <div id="div_bottom" runat="server"></div>
    </form>
</body>
</html>
