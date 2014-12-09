<%@ Control Language="VB" ClassName="LogonPanel" %>
<%@ import  Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '         'WriteDebug("StyleSheetPath = " & Session("StyleSheetPath"))          ' CN

    Const USER_PERMISSION_VIEW_STOCK As Integer = 1024
    Const USER_PERMISSION_CREATE_STOCK_BOOKING As Integer = 2048
    Const USER_PERMISSION_PRINT_ON_DEMAND_TAB As Integer = 4096
    Const USER_PERMISSION_ADVANCED_PERMISSIONS_TAB As Integer = 8192
    Const USER_PERMISSION_FILE_UPLOAD_TAB As Integer = 16384

    Const CUSTOMER_WURS As Int32 = 579
    Const CUSTOMER_WUIRE As Int32 = 686
    Const CUSTOMER_WURSDEMO As Int32 = 788

    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Public Event evMustChangePassword()
    Public Event evPasswordExpired()
    
    Dim sUserID As String
    Dim sPassword As String
    Dim bAutoLogon As Boolean = False
    Private bMustChangePassword As Boolean
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gbIsAutologon As Boolean = False
    Private bPrintOnDemandTab As Boolean
    Private bUserPermissionsTab As Boolean
    Private bFileUploadTab As Boolean
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.SetFocus(txtUserId)
        sUserID = txtUserId.Text
        sPassword = txtPassword.Text
        Call GetSiteKeyFromSiteName(sGetPath)
        If bIsKnownUser() Then
            rfvPassword.EnableClientScript = False
            rfvPassword.Enabled = False
        End If
        If Request.Cookies("SprintLogon") IsNot Nothing Then
            If Request.Cookies("SprintLogon")("UserID") <> String.Empty Then
                txtUserId.Text = Request.Cookies("SprintLogon")("UserID")
                txtPassword.Text = Request.Cookies("SprintLogon")("Password")
                If txtUserId.Text.EndsWith(",") Then
                    txtUserId.Text = String.Empty
                    txtPassword.Text = String.Empty
                Else
                    gbIsAutologon = True
                    If Not pbAutoLogonFailed Then
                        Call DoLogon()
                    End If
                End If
            End If
        End If
        txtUserId.Attributes.Add("onkeypress", "return clickButton(event,'" + btnLogon.ClientID + "')")
        txtPassword.Attributes.Add("onkeypress", "return clickButton(event,'" + btnLogon.ClientID + "')")
    End Sub

    Protected Sub WriteDebug(ByVal sString As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "INSERT INTO AAA_Debug (result) VALUES ('" & sString.Replace("'", "''") & "')"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("WriteDebug: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    'Protected Function sGetPath() As String
    '    Dim sPathInfo As String = Request.Path
    '    sGetPath = String.Empty
    '    If sPathInfo <> String.Empty Then
    '        sPathInfo = sPathInfo.Substring(1)
    '        Dim sPos As Integer = sPathInfo.IndexOf("/")
    '        If sPos > 0 Then
    '            sGetPath = sPathInfo.Substring(0, sPos)
    '        End If
    '    End If
    'End Function

    Protected Function sGetPath() As String
        sGetPath = String.Empty
        If Request.Url.ToString.ToLower.Contains("jupitermarketing") Then
            sGetPath = "jupiter"
        Else
            'WriteDebug("URL = " & Request.Url.ToString)          ' CN
            Dim sPathInfo As String = Request.Path.Trim
            'WriteDebug("sPathInfo = |" & sPathInfo & "|")          ' CN
            If sPathInfo.Replace("/default.aspx", "") <> String.Empty Then
                sGetPath = sPathInfo.Substring(1)
                sGetPath = sGetPath.Replace("/default.aspx", "")
            Else
                sGetPath = Request.Url.ToString.Substring(7)
                sGetPath = sGetPath.Replace("/default.aspx", "")
                sGetPath = sGetPath.Replace("www.", "")
                sGetPath = sGetPath.Replace("www2.", "")
                sGetPath = sGetPath.Replace(".co.uk", "")
                sGetPath = sGetPath.Replace(".com", "")
                sGetPath = sGetPath.Replace(".net", "")
            End If
            'WriteDebug("sGetPath = " & sGetPath)          ' CN
        End If
    End Function
    
    Protected Sub btnLogon_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call DoLogon()
    End Sub
    
    Protected Function bIsKnownUser() As Boolean
        bIsKnownUser = False
        Dim sKnownAddresses() As String = {"127.0.0.1", "81.138.76.105", "81.168.63.93", "82.152.142.10", "135.196.230.102", "217.205.106.17", "217.205.106.18", "86.188.157.138", "87.74.42.54", "::1", "86.188.223.114"}
        Dim sUserHostAddr As String = Request.UserHostAddress
        If Array.IndexOf(sKnownAddresses, sUserHostAddr) >= 0 Or sUserHostAddr.StartsWith("10.") Or sUserHostAddr.StartsWith("11.") Then
            bIsKnownUser = True
        End If
    End Function
    
    Protected Sub DoLogon()
        If (Len(txtUserId.Text) > 0 And (Len(txtPassword.Text) > 0 Or bIsKnownUser())) Then
            Dim oUserInfo As Transworld.UserInfo = New Transworld.UserInfo()
            Dim oLogon As Transworld.Logon = New Transworld.Logon()
            Dim oPassword As Transworld.Password = New Transworld.Password()
            Dim sUserURL As String
            pbAutoLogonFailed = True  ' posit failure
    
            oUserInfo = oLogon.GetUserInfo(txtUserId.Text)
    
            If oUserInfo.UserKey = -1 Or ((oPassword.Decrypt(oUserInfo.Password) <> txtPassword.Text) And Not bIsKnownUser()) Then
                lblErrorMessage.Text = "User ID or password unrecognised"
                Call RemoveAutoLogonCookie()
                If txtUserId.Text <> psUserId Then
                    pnRetries = 0
                    psUserId = txtUserId.Text
                Else
                    pnRetries = pnRetries + 1
                End If
                If pnRetries >= oUserInfo.MaxPasswordRetries Then
                    Call SetAccountSuspended()
                    lblErrorMessage.Text = "Max login attempts exceeded - account suspended"
                    pnRetries = 0
                End If
                Exit Sub
            End If
            
            If oUserInfo.Status = "Suspended" Then
                lblErrorMessage.Text = "Account suspended"
                Call RemoveAutoLogonCookie()
                Exit Sub
            End If
                
            If oUserInfo.AccountDisabledDueToInactivity Then
                lblErrorMessage.Text = "Account disabled due to inactivity"
                Call RemoveAutoLogonCookie()
                Exit Sub
            End If

            pbAutoLogonFailed = False  ' indicate success

            If cbRememberMe.Checked Then
                Dim c As HttpCookie
                If (Request.Cookies("SprintLogon") Is Nothing) Then
                    c = New HttpCookie("SprintLogon")
                Else
                    c = Request.Cookies("SprintLogon")
                End If
                c.Values.Add("UserID", txtUserId.Text)
                c.Values.Add("Password", txtPassword.Text)
                c.Expires = DateTime.Now.AddDays(7)
                Response.Cookies.Add(c)
            End If
                
            Session("UserKey") = oUserInfo.UserKey
            Session("CustomerKey") = oUserInfo.CustomerKey
            Session("CustomerName") = oUserInfo.CustomerName
            Session("UserName") = oUserInfo.UserName
            Session("UserType") = oUserInfo.UserType
            If Not (oUserInfo.RunningHeaderImage = "default") Or (oUserInfo.RunningHeaderImage = String.Empty) Then
                Session("RunningHeaderImage") = oUserInfo.RunningHeaderImage
            End If
            Session("ViewGAB") = oUserInfo.AbleToViewGlobalAddressBook
            Session("EditGAB") = oUserInfo.AbleToEditGlobalAddressBook
            'Session("AbleToCreateStockBooking") = oUserInfo.AbleToCreateStockBooking
            Session("AbleToCreateCollectionRequest") = oUserInfo.AbleToCreateCollectionRequest
            Session("ApplyStockMaxGrabRule") = oUserInfo.ApplyStockMaxGrabRule
            If oUserInfo.AbleToViewStock Then
                oUserInfo.UserPermissions += USER_PERMISSION_VIEW_STOCK
            End If
            If oUserInfo.AbleToCreateStockBooking Then
                oUserInfo.UserPermissions += USER_PERMISSION_CREATE_STOCK_BOOKING
            End If
            If bPrintOnDemandTab Then
                oUserInfo.UserPermissions += USER_PERMISSION_PRINT_ON_DEMAND_TAB
            End If
            
            If bUserPermissionsTab Then
                oUserInfo.UserPermissions += USER_PERMISSION_ADVANCED_PERMISSIONS_TAB
            End If
            
            If bFileUploadTab Then
                oUserInfo.UserPermissions += USER_PERMISSION_FILE_UPLOAD_TAB
            End If
            
            Session("UserPermissions") = oUserInfo.UserPermissions
            Session("LastLogon") = oUserInfo.LastLogon

            Select Case oUserInfo.UserType.ToLower
                Case "sa".ToLower
                    'sUserURL = ConfigLib.GetConfigItem_SA_Home_Page
                    Session.Timeout = 30
                Case "SuperUser".ToLower
                    'sUserURL = ConfigLib.GetConfigItem_SuperUser_Page
                    Session.Timeout = 20
                Case "ProductOwner".ToLower
                    'sUserURL = ConfigLib.GetConfigItem_ProductOwner_Home_Page
                    Session.Timeout = 20
                Case "User".ToLower
                    'sUserURL = ConfigLib.GetConfigItem_User_Home_Page
                    Session.Timeout = 20
                Case Else
                    'sUserURL = ConfigLib.GetConfigItem_User_Home_Page
                    Session.Timeout = 20
            End Select
            Session("CustomerCreatedOn") = oUserInfo.CustomerCreatedOn

            If UsesNewRotator() Or oUserInfo.CustomerCreatedOn > DateTime.Parse("11-Jul-2012") Then
                sUserURL = "NoticeBoard2.aspx"
            Else
                sUserURL = "NoticeBoard.aspx"
            End If
            sUserURL = "NoticeBoard.aspx"   ' temporarily CN 25OCT12
            
            Call UpdateLastLogonTime()

            Dim tsPasswordExpiryDays As TimeSpan = TimeSpan.FromDays(oUserInfo.PasswordExpiryDays)
            Dim tsNextPasswordChange As DateTime = oUserInfo.LastPasswordChange + tsPasswordExpiryDays
            If tsNextPasswordChange < DateTime.Now Then
                If gbIsAutologon Then
                    Call RemoveAutoLogonCookie()
                    Exit Sub
                Else
                    RaiseEvent evPasswordExpired()
                    Exit Sub
                End If
            End If
            
            If oUserInfo.MustChangePassword Then
                If gbIsAutologon Then
                    Call RemoveAutoLogonCookie()
                    Exit Sub
                Else
                    RaiseEvent evMustChangePassword()
                    Exit Sub
                End If
            End If
            Call SetStyleSheetPath()
            Response.Redirect(sUserURL)    ' was Server.Transfer(sUserURL)
        End If
    End Sub
    
    Protected Function UsesNewRotator() As Boolean
        Dim arrUsesNewRotator() As Integer = {CUSTOMER_WURS, CUSTOMER_WUIRE, CUSTOMER_WURSDEMO}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        UsesNewRotator = IIf(gbSiteTypeDefined, gsSiteType = "newrotator", Array.IndexOf(arrUsesNewRotator, nCustomerKey) >= 0)
    End Function

    Protected Sub SetStyleSheetPath()
        Const DEFAULT_STYLESHEET_PATH As String = ".\css\sprint.css"
        Dim sStyleSheetPath As String = DEFAULT_STYLESHEET_PATH
        Dim sPathInfo As String = sGetPath()
        If sPathInfo <> String.Empty Then

            If sPathInfo.Contains("?") Then
                Dim nStartPos As Integer = sPathInfo.IndexOf("?")
                sPathInfo = sPathInfo.Substring(0, nStartPos)
            End If
            
            sPathInfo = ".\css\sprint_.css "
            If My.Computer.FileSystem.FileExists(Request.MapPath(sPathInfo)) Then
                sStyleSheetPath = sPathInfo
            End If
        End If
        Session("StyleSheetPath") = sStyleSheetPath
    End Sub
    
    Protected Function GetSiteKeyFromSiteName(ByVal sSiteName As String) As Integer
        GetSiteKeyFromSiteName = 0
        If sSiteName <> String.Empty Then
            Dim oDataReader As SqlDataReader = Nothing
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Site_MapNameToKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
        
            Dim oParamSiteName As SqlParameter = oCmd.Parameters.Add("@SiteName", SqlDbType.VarChar, 50)
            oParamSiteName.Value = sSiteName
        
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                If oDataReader.HasRows Then
                    oDataReader.Read()
                    GetSiteKeyFromSiteName = oDataReader("SiteKey")
                    Session("SiteKey") = GetSiteKeyFromSiteName
                Else
                    Session("SiteKey") = 0
                End If
            Catch ex As Exception
                WebMsgBox.Show("GetSiteKeyFromSiteName: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        Else
            Session("SiteKey") = 0
        End If
        Call GetSiteSettings()
    End Function
    
    Protected Sub GetSiteSettings()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent3", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettings"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteSettings: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        If oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            Session("SiteTitle") = dr("SiteTitle") & String.Empty
            Page.Title = Session("SiteTitle")
            Session("RunningHeaderImage") = dr("DefaultRunningHeaderImage") & String.Empty
            bPrintOnDemandTab = CBool(dr("Misc1"))
            bUserPermissionsTab = CBool(dr("UserPermissions"))
            bFileUploadTab = CBool(dr("FileUpload"))
        Else
            Session("SiteTitle") = String.Empty
            Session("RunningHeaderImage") = String.Empty
        End If
    End Sub
    
    Protected Sub SetAccountSuspended()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "UPDATE UserProfile SET Status = 'Suspended' WHERE UserId = '" & psUserId & "'"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SetAccountSuspended: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ' send email to administrator
    End Sub
    
    Protected Sub SendEmail()
        
    End Sub
    
    Protected Sub UpdateLastLogonTime()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_UpdateLastLogon", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = CLng(Session("UserKey"))
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("UpdateLastLogonTime: " & ex.Message)
        Finally
            oConn.Close()
        End Try
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
    
    Public WriteOnly Property Logon() As System.Delegate
      Set (ByVal Value As System.Delegate)
    '    _delLogon = Value
      End Set
    End Property

    Protected Sub SetAutoLogon()
        Dim c As HttpCookie = New HttpCookie("SprintLogon")
        c.Values.Add("UserID", sUserID)
        c.Values.Add("Password", sPassword)
        c.Expires = DateTime.Now.AddDays(7)
        Response.Cookies.Add(c)
        Response.Flush()
    End Sub
    
    Public Property pbMustChangePassword() As Boolean
        Get
            Return bMustChangePassword
        End Get
        Set(ByVal value As Boolean)
            bMustChangePassword = value
        End Set
    End Property
    
    Property pbAutoLogonFailed() As Boolean
        Get
            Dim o As Object = ViewState("LP_AutoLogonFailed")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("LP_AutoLogonFailed") = Value
        End Set
    End Property
        
    Property pnRetries() As Integer
        Get
            Dim o As Object = ViewState("LP_Retries")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("LP_Retries") = Value
        End Set
    End Property
   
    Property psUserId() As String
        Get
            Dim o As Object = ViewState("LP_UserId")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("LP_UserId") = Value
        End Set
    End Property
   
</script>
<style type="text/css">
A            {color:#000000; font-size:8pt; font-family: arial}
A:LINK        {Text-Decoration: none; color:#000000; font-size:8pt; font-family: arial}
A:VISITED    {Text-Decoration: none; color:#000000; font-size:8pt; font-family: arial}
A:HOVER        {Text-Decoration: none; color:#000000; font-size:8pt; font-family: arial}
</style>
<table style="width:330px; border-right: DarkGray 1px solid; border-top: dimgray 1px solid; border-left: dimgray 1px solid; border-bottom: dimgray 1px solid; font-family:Arial; background-color: #f0f0f0; font-size: x-small;"  >
    <tr>
        <td style="width: 64px">
            &nbsp;
        </td>
        <td style="width: 80%">
        </td>
    </tr>
    <tr>
        <td style="width: 64px">
            <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">&nbsp;User ID </asp:Label>
        </td>
        <td>
            <asp:TextBox runat="server" TabIndex="1" Width="180px" ID="txtUserId" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
            <asp:RequiredFieldValidator ID="rfvUserId" runat="server" ControlToValidate="txtUserId" Font-Names="Verdana" Font-Size="XX-Small" EnableClientScript="False">required</asp:RequiredFieldValidator></td>
    </tr>
    <tr>
        <td style="width: 64px">
            <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">&nbsp;Password</asp:Label>
        </td>
        <td>
            <asp:TextBox runat="server" TextMode="Password" TabIndex="2" Width="180px" ID="txtPassword" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
            <asp:RequiredFieldValidator ID="rfvPassword" runat="server" ControlToValidate="txtPassword" Font-Names="Verdana" Font-Size="XX-Small" EnableClientScript="False">required</asp:RequiredFieldValidator></td>
    </tr>
    <tr>
        <td style="width: 64px; height: 26px;">
        </td>
        <td style="height: 26px">
            <asp:Button runat="server" TabIndex="3" ID="btnLogon" OnClick="btnLogon_Click" Text="log in"></asp:Button>
            &nbsp; &nbsp;
            <asp:Label runat="server" ForeColor="Red" ID="lblErrorMessage" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
        </td>
    </tr>
    <tr>
        <td style="width: 64px" align="center" valign="middle">
            <asp:HyperLink ID="hlnkForgotPassword" runat="server" Font-Names="Arial" Font-Size="XX-Small"
                ForeColor="Blue" NavigateUrl="RequestCredentials.aspx">forgot your<br />password?</asp:HyperLink></td>
        <td>
            <asp:CheckBox ID="cbRememberMe" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                Text="remember me on this computer" /><br />
            &nbsp;<asp:Label ID="Label3" runat="server" Font-Size="XX-Small" Font-Names="Verdana" Font-Italic="True">Click </asp:Label>
            <asp:HyperLink runat="server" NavigateUrl="request_userid.aspx?type=newuser" ForeColor="Blue"
                ID="HyperLink1" Target="_blank" Font-Size="XX-Small" Font-Names="Verdana" Font-Italic="True">here</asp:HyperLink>
            <asp:Label ID="Label4" runat="server" Font-Size="XX-Small" Font-Names="Verdana" Font-Italic="True"> if you require access to the system</asp:Label></td>
    </tr>
</table>