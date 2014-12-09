<%@ Page Language="VB" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header_v2.ascx" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' TO DO
    ' Sort out table width change when Hidden selected
    
    Dim sConn As String = System.Configuration.ConfigurationManager.AppSettings("ConnectionString")
    Dim oConn As New SqlConnection(sConn)
    Dim oCmd As New SqlCommand("spASPNET_Customer_UserContent", oConn)
    Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_UserContent", oConn)
    Dim oDataSet As New DataSet()

    Sub Page_Load(ByVal sender As System.Object, ByVal e As EventArgs)

        Rotator1.Visible = True
        Rotator1.ScrollDirection = ScrollDirection.Left

        oConn.Open()
    
        If Not IsPostBack Then
            rblHomePageRotatorPreset.DataBind()
            rblHeaderRotatorPreset.DataBind()
            Try                                                 ' check an entry exists for this customer, create one if not
                oCmd.CommandType = CommandType.StoredProcedure
                
                oCmd.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
                oCmd.Parameters("@Action").Value = "VERIFY"
                oCmd.Parameters("@Action").Direction = ParameterDirection.Input
                
                oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
                oCmd.Parameters("@CustomerKey").Direction = ParameterDirection.Input
                
                oCmd.ExecuteNonQuery()
            
            Catch ex As SqlException
                ' do something to notify user here !!!!!!!!!!!!!!
            End Try
            InitFromUserContent()
            Call ShowPreset(rblHomePageRotatorPreset.SelectedIndex, Rotator1)
            Call ShowPreset(rblHeaderRotatorPreset.SelectedIndex, Rotator2)
        Else
            Call WriteUserContent()
        End If
    End Sub
    
    Sub InitFromUserContent()
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        oAdapter.SelectCommand.Parameters("@Action").Direction = ParameterDirection.Input
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.Parameters("@CustomerKey").Direction = ParameterDirection.Input
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Direction = ParameterDirection.Input
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "Rotators"
        
        oAdapter.Fill(oDataSet, "Rotators")
        Dim nHeaderPagePreset As Integer
        Dim nHomePagePreset As Integer
        nHeaderPagePreset = CInt(oDataSet.Tables(0).Rows(0).Item(0))
        nHomePagePreset = CInt(oDataSet.Tables(0).Rows(0).Item(1))
        rblHeaderRotatorPreset.SelectedIndex = nHeaderPagePreset
        rblHomePageRotatorPreset.SelectedIndex = nHomePagePreset
    End Sub 'InitFromUserContent
    
    Sub WriteUserContent()
        Try
            oCmd.CommandType = CommandType.StoredProcedure
            
            oCmd.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Action").Value = "SET"
            oCmd.Parameters("@Action").Direction = ParameterDirection.Input
            
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oCmd.Parameters("@CustomerKey").Direction = ParameterDirection.Input
            
            oCmd.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@ContentType").Value = "Rotators"
            oCmd.Parameters("@ContentType").Direction = ParameterDirection.Input
            
            oCmd.Parameters.Add(New SqlParameter("@HeaderRotatorPreset", SqlDbType.Int))
            oCmd.Parameters("@HeaderRotatorPreset").Value = rblHeaderRotatorPreset.SelectedIndex
            oCmd.Parameters("@HeaderRotatorPreset").Direction = ParameterDirection.Input
            
            oCmd.Parameters.Add(New SqlParameter("@HomePageRotatorPreset", SqlDbType.Int))
            oCmd.Parameters("@HomePageRotatorPreset").Value = rblHomePageRotatorPreset.SelectedIndex
            oCmd.Parameters("@HomePageRotatorPreset").Direction = ParameterDirection.Input
            
            oCmd.ExecuteNonQuery()
            
        Catch ex As SqlException
            ' do something to notify user here !!!!!!!!!!!!!!
        End Try
        
    End Sub
    
    Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As EventArgs) ' does a round trip to save current values    
    End Sub

    Function RotatorPresets() As String()
        Return New String() {"hidden", "no animation", "scroll up", "scroll left", "slow fade in/out", "normal fade in/out", "fast fade in/out", "slow pixelate in/out", "normal pixelate in/out", "fast pixelate in/out", "slow dissolve in/out", "normal dissolve in/out", "fast dissolve in/out", "slow gradient wipe in/out", "normal gradient wipe in/out", "fast gradient wipe in/out"}
    End Function
    
    Protected Function GetSelectedPreset(ByVal preset As Object) As Integer
        Return Array.IndexOf(RotatorPresets, preset.ToString())
    End Function
    
    Sub ShowPreset(ByVal nPreset As Integer, ByVal Rotator As ComponentArt.Web.UI.Rotator)
        Rotator.Visible = True
        Rotator.RotationType = RotationType.SlideShow
        Select Case nPreset
            Case 0              ' hidden
                Rotator.Visible = False
            Case 1              ' no animation
                Rotator.ShowEffect = RotationEffect.None
            Case 2              ' scroll up
                Rotator.RotationType = RotationType.SmoothScroll
                Rotator.ScrollDirection = ScrollDirection.Up
                Rotator.ScrollInterval = 15
            Case 3              ' scroll left
                Rotator.RotationType = RotationType.SmoothScroll
                Rotator.ScrollDirection = ScrollDirection.Left
                Rotator.ScrollInterval = 15
            Case 4              ' slow fade in/out
                Rotator.ShowEffect = RotationEffect.Fade
                Rotator.ShowEffectDuration = 2500
                Rotator.HideEffect = RotationEffect.Fade
                Rotator.HideEffectDuration = 2500
            Case 5              ' normal fade in/out
                Rotator.ShowEffect = RotationEffect.Fade
                Rotator.ShowEffectDuration = 1500
                Rotator.HideEffect = RotationEffect.Fade
                Rotator.HideEffectDuration = 1500
            Case 6              ' fast fade in/out
                Rotator.ShowEffect = RotationEffect.Fade
                Rotator.ShowEffectDuration = 500
                Rotator.HideEffect = RotationEffect.Fade
                Rotator.HideEffectDuration = 500
            Case 7              ' slow pixelate in/out
                Rotator.ShowEffect = RotationEffect.Pixelate
                Rotator.ShowEffectDuration = 2500
                Rotator.HideEffect = RotationEffect.Pixelate
                Rotator.HideEffectDuration = 2500
            Case 8              ' normal pixelate in/out
                Rotator.ShowEffect = RotationEffect.Pixelate
                Rotator.ShowEffectDuration = 1500
                Rotator.HideEffect = RotationEffect.Pixelate
                Rotator.HideEffectDuration = 1500
            Case 9              ' fast pixelate in/out
                Rotator.ShowEffect = RotationEffect.Pixelate
                Rotator.ShowEffectDuration = 500
                Rotator.HideEffect = RotationEffect.Pixelate
                Rotator.HideEffectDuration = 500
            Case 10              ' slow dissolve in/out
                Rotator.ShowEffect = RotationEffect.Dissolve
                Rotator.ShowEffectDuration = 2500
                Rotator.HideEffect = RotationEffect.Dissolve
                Rotator.HideEffectDuration = 2500
            Case 11              ' normal dissolve in/out
                Rotator.ShowEffect = RotationEffect.Dissolve
                Rotator.ShowEffectDuration = 1500
                Rotator.HideEffect = RotationEffect.Dissolve
                Rotator.HideEffectDuration = 1500
            Case 12              ' fast dissolve in/out
                Rotator.ShowEffect = RotationEffect.Dissolve
                Rotator.ShowEffectDuration = 500
                Rotator.HideEffect = RotationEffect.Dissolve
                Rotator.HideEffectDuration = 500
            Case 13              ' slow gradient wipe in/out
                Rotator.ShowEffect = RotationEffect.GradientWipe
                Rotator.ShowEffectDuration = 2500
                Rotator.HideEffect = RotationEffect.GradientWipe
                Rotator.HideEffectDuration = 2500
            Case 14              ' normal gradient wipe in/out
                Rotator.ShowEffect = RotationEffect.GradientWipe
                Rotator.ShowEffectDuration = 1500
                Rotator.HideEffect = RotationEffect.GradientWipe
                Rotator.HideEffectDuration = 1500
            Case 15              ' fast gradient wipe in/out
                Rotator.ShowEffect = RotationEffect.GradientWipe
                Rotator.ShowEffectDuration = 500
                Rotator.HideEffect = RotationEffect.GradientWipe
                Rotator.HideEffectDuration = 500
        End Select
    End Sub
    
    Protected Sub rblHomePageRotatorPreset_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowPreset(rblHomePageRotatorPreset.SelectedIndex, Rotator1)
        Call ShowPreset(rblHeaderRotatorPreset.SelectedIndex, Rotator2)
    End Sub

    Protected Sub rblHeaderRotatorPreset_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowPreset(rblHomePageRotatorPreset.SelectedIndex, Rotator1)
        Call ShowPreset(rblHeaderRotatorPreset.SelectedIndex, Rotator2)
    End Sub
</script>
<html>
<head>
    <link href="elog.css" type="text/css" rel="stylesheet" />
    <link href="tabs.css" rel="STYLESHEET" type="text/css" />
    <style type="text/css" media="screen">
BODY {
	FONT-FAMILY: sans-serif
}
TABLE {
	FONT-SIZE: 7pt; FONT-FAMILY: verdana
}
TD.small {
	FONT-SIZE: 8pt; FONT-FAMILY: verdana
}
TD.subheading {
	FONT-SIZE: 14pt; FONT-FAMILY: sans-serif
}
TR.darkbackground {
	BACKGROUND-COLOR: silver
}
</style>
    
</head>
<body class="sf">
    <form runat="server">
        <p>
            <main:Header id="ctlHeader" runat="server"></main:Header>
        </p>
                <p></p>
        <table border="0" cellpadding="0" cellspacing="0" style="width: 95%">
            <tr>
                <td align="center" style="height: 12px"><b>Home Page Animation</b>
</td>
                <td align="center" style="height: 12px">
                </td>
                <td align="center" style="height: 12px"><b>Top of Page Animation</b>
</td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td>
                    &nbsp; &nbsp; &nbsp;
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td align="center">
                <COMPONENTART:ROTATOR id="Rotator1" runat="server" width="400" height="50" CssClass="Rotator" XmlContentFile="home_page.xml">
                    <SLIDETEMPLATE>
                        <table cellspacing="1" cellpadding="0" width="100%" border="0" bgcolor="#eeeeee">
                            <tr>
                                <td class="RotatorMain">
                                    <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    <img src='images/rotatorExampleImage.jpg' height="44" /></span>
                                </td>
                                <td class="RotatorMain" nowrap="nowrap">
                                    <span id="RotatorText" runat="server" class="AdRotatorText">
                                        This is how the home page<br /> animation will appear </span>
                                </td>
                            </tr>
                        </table>
                    </SLIDETEMPLATE>
                </COMPONENTART:ROTATOR>
                </td>
                <td>
                </td>
                <td align="center">
                <COMPONENTART:ROTATOR id="Rotator2" runat="server" width="400" height="50" CssClass="Rotator" XmlContentFile="home_page.xml">
                    <SLIDETEMPLATE>
                        <table cellspacing="1" cellpadding="0" width="100%" border="0" bgcolor="#eeeeee">
                            <tr>
                                <td class="RotatorMain">
                                    <span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    <img src="images/rotatorExampleImage.jpg" height="44" /></span>
                                </td>
                                <td class="RotatorMain" nowrap="nowrap">
                                    <span id="RotatorText" runat="server" class="AdRotatorText">
                                        This is how the top of page <br /> animation will appear </span>
                                </td>
                            </tr>
                        </table>
                    </SLIDETEMPLATE>
                </COMPONENTART:ROTATOR>
                </td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td align="center">
                    <asp:RadioButtonList
                            ID="rblHomePageRotatorPreset"
                            Datasource="<%# RotatorPresets() %>"
                            runat="server"
                            AutoPostBack="True"
                            RepeatDirection="Horizontal" RepeatColumns="2"
                            OnSelectedIndexChanged="rblHomePageRotatorPreset_SelectedIndexChanged">
                    </asp:RadioButtonList></td>
                <td align="center">
                </td>
                <td align="center">
                    <asp:RadioButtonList
                            ID="rblHeaderRotatorPreset"
                            Datasource="<%# RotatorPresets() %>"
                            runat="server"
                            AutoPostBack="True"
                            RepeatDirection="Horizontal" RepeatColumns="2"
                            OnSelectedIndexChanged="rblHeaderRotatorPreset_SelectedIndexChanged">
                    </asp:RadioButtonList></td>
            </tr>
        </table>
    </form>
</body>
</html>
