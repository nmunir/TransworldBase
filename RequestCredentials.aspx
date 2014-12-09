<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<script runat="server">
    Dim sSprintReceipient As String = "account.managers@transworld.eu.com"
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim sSentFromEmail As String = "automailer@transworld.eu.com"

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("UserKey")) And pbLoggedOnAtSessionStart = True Then
            Server.Transfer("session_expired.aspx")
        End If

        If Not IsPostBack Then
            If IsNothing(Session("SiteKey")) Then              ' is nothing on first load as panel has not yet been loaded to init SiteKey
                Call GetSiteKeyFromSiteName(sGetPath)
            End If
            Call GetSiteSettings()
            If Session("RunningHeaderImage") = String.Empty Then
                imgLogo.Visible = False
            Else
                imgLogo.ImageUrl = Session("RunningHeaderImage")
            End If
            Call HideAllPanels()
            If IsNumeric(Session("UserKey")) Then
                Page.Title = "Change Password"
                pbLoggedOnAtSessionStart = True
                pnlRequestLogonDetails.Visible = False
            Else
                pnlRequestLogonDetails.Visible = True
                tbRecoveryEmailAddr.Focus()
            End If
        End If
        tbRecoveryEmailAddr.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitRecoveryRequest.ClientID + "')")
        'imgLogo.ImageUrl = ConfigLib.GetConfigItem_Default_Running_Header_Image
        Call SetTitle()
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Request Credentials"
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
    End Function
    
    Protected Function sGetPath() As String
        Dim sPathInfo As String = Request.Path
        sGetPath = String.Empty
        If sPathInfo <> String.Empty Then
            sPathInfo = sPathInfo.Substring(1)
            Dim sPos As Integer = sPathInfo.IndexOf("/")
            If sPos > 0 Then
                sGetPath = sPathInfo.Substring(0, sPos)
            End If
        End If
    End Function
    
    Protected Sub GetSiteSettings()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent", oConn)
        
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
            Session("RunningHeaderImage") = dr("DefaultRunningHeaderImage") & String.Empty
        Else
            Session("SiteTitle") = String.Empty
            Session("RunningHeaderImage") = String.Empty
        End If
    End Sub

    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub HideAllPanels()
        pnlRequestLogonDetails.Visible = False
        pnlCloseWindow.Visible = False
        pnlSelectUserId.Visible = False
    End Sub
    
    Protected Sub ShowSelectUserIdPanel()
        Call HideAllPanels()
        pnlSelectUserId.Visible = True
    End Sub
    
    Protected Sub SendEmail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBody As String)
        Dim bError As Boolean = False

        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_AddEmailToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageTypeId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageTypeId").Value = "WEB_USERID_REQUEST"
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            If IsNumeric(Session("UserKey")) Then
                oCmd.Parameters("@CustomerKey").Value = Session("UserKey")
            Else
                oCmd.Parameters("@CustomerKey").Value = 0
            End If
    
            oCmd.Parameters.Add(New SqlParameter("@Recipient", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Recipient").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@Body", SqlDbType.NVarChar, 3880))
            oCmd.Parameters("@Body").Value = sBody
    
            oCmd.Parameters.Add(New SqlParameter("@EmailMessageQueueKey", SqlDbType.Int, 4))
            oCmd.Parameters("@EmailMessageQueueKey").Direction = ParameterDirection.Output
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SendEmail: " & ex.Message)
            bError = True
        Finally
            oConn.Close()
        End Try
    End Sub

    <Serializable()> Public Class Memorable
        Public MemorableQuestion As String
        Public MemorableAnswer As String
    End Class
    
    Protected Sub RetrieveCurrentLogonDetails()
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey2", oConn)
        
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = pnUserKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("UserId")) Then
                psCurrentUserId = oPassword.Decrypt(oDataReader("UserId"))
            End If
            If Not IsDBNull(oDataReader("Password")) Then
                psCurrentPassword = oPassword.Decrypt(oDataReader("Password"))
            End If
            Dim alMemorableAnswers As New ArrayList
            If Not IsDBNull(oDataReader("MemorableAnswer1")) AndAlso oDataReader("MemorableAnswer1").ToString.Length > 0 Then
                Dim oMemorable As Memorable = New Memorable
                oMemorable.MemorableQuestion= "Birthday of your spouse / partner / significant other (dd/mm/yy)?"
                oMemorable.MemorableAnswer = oDataReader("MemorableAnswer1")
                alMemorableAnswers.Add(oMemorable)
            End If
            If Not IsDBNull(oDataReader("MemorableAnswer2")) AndAlso oDataReader("MemorableAnswer2").ToString.Length > 0 Then
                Dim oMemorable As Memorable = New Memorable
                oMemorable.MemorableQuestion = "Street number of the house you first lived in?"
                oMemorable.MemorableAnswer = oDataReader("MemorableAnswer2")
                alMemorableAnswers.Add(oMemorable)
            End If
            If Not IsDBNull(oDataReader("MemorableAnswer3")) AndAlso oDataReader("MemorableAnswer3").ToString.Length > 0 Then
                Dim oMemorable As Memorable = New Memorable
                oMemorable.MemorableQuestion = "Grandmother's maiden name?"
                oMemorable.MemorableAnswer = oDataReader("MemorableAnswer3")
                alMemorableAnswers.Add(oMemorable)
            End If
            If alMemorableAnswers.Count > 0 Then
                palMemorable = alMemorableAnswers
            End If

        Catch ex As Exception
            WebMsgBox.Show("RetrieveCurrentLogonDetails: " & ex.Message)
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetMatchingEmailAddresses() As System.Collections.Generic.Dictionary(Of String, String)
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim dictAccounts As New System.Collections.Generic.Dictionary(Of String, String)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfilesbyEmailAddr", oConn)
        
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@EmailAddr", SqlDbType.NVarChar, 100)
        oParam.Value = tbRecoveryEmailAddr.Text
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                Do While oDataReader.Read()
                    Dim sUserId As String = String.Empty, sPassword As String = String.Empty
                    If Not IsDBNull(oDataReader("UserId")) Then
                        sUserId = oDataReader("UserId")
                    End If
                    If Not IsDBNull(oDataReader("Password")) Then
                        sPassword = oPassword.Decrypt(oDataReader("Password"))
                    End If
                    dictAccounts.Add(sUserId, sPassword)
                    Dim li As New ListItem(oDataReader("UserId"), oDataReader("Key"))
                    lbAccounts.Items.Add(li)
                Loop
            Else
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetMatchingEmailAddresses: " & ex.Message)
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
        GetMatchingEmailAddresses = dictAccounts
    End Function

    Protected Sub btnSubmitRecoveryRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SubmitRecoveryRequest(0)
    End Sub

    'Public Sub New(ByVal SI As Runtime.Serialization.SerializationInfo, ByVal SC As Runtime.Serialization.StreamingContext)
    '    MyBase.New(SI, SC)
    'End Sub
    
    Protected Sub SubmitRecoveryRequest(ByVal nUserKey As Integer)
        'Dim dictAccounts As System.Collections.Generic.Dictionary(Of String, String)
        Page.Validate()
        If Page.IsValid Then
            If Not trMemorableQuestion.Visible Then
                pdictAccounts = GetMatchingEmailAddresses()
                If pdictAccounts.Count = 0 Then
                    WebMsgBox.Show("No account matches this email address")
                Else
                    If nUserKey > 0 Then
                        pnUserKey = nUserKey
                        trMemorableQuestion.Visible = True
                        tbRecoveryEmailAddr.Enabled = False
                    Else
                        If pdictAccounts.Count = 1 Then
                            pnUserKey = lbAccounts.Items(0).Value
                            trMemorableQuestion.Visible = True
                            tbRecoveryEmailAddr.Enabled = False
                        Else
                            Call ShowSelectUserIdPanel()
                            Exit Sub
                        End If
                    End If
                    Call RetrieveCurrentLogonDetails()
                    If palMemorable Is Nothing OrElse palMemorable.Count = 0 Then
                        trMemorableQuestion.Visible = False
                        tbRecoveryEmailAddr.Enabled = True
                        tbRecoveryEmailAddr.Text = ""
                        WebMsgBox.Show("You have not entered any security answers. Cannot email access details. Please contact customer services.")
                        Exit Sub
                    Else
                        Dim RandomClass As New Random
                        Dim oMemorable As Memorable = palMemorable(RandomClass.Next(0, palMemorable.Count))
                        lblMemorableQuestion.Text = oMemorable.MemorableQuestion.ToString
                        psExpectedMemorableAnswer = oMemorable.MemorableAnswer.ToString
                    End If
                End If
            Else
                If tbMemorableAnswer.Text.Length > 0 Then
                    If tbMemorableAnswer.Text = psExpectedMemorableAnswer Then
                        Call SendEmail("WEB_LOGON_REQUEST", tbRecoveryEmailAddr.Text, "Request for Transworld logon details", sBuildLogonDetailsMessageBody(pdictAccounts))
                        WebMsgBox.Show("An email containing your logon details has been sent to " & tbRecoveryEmailAddr.Text & " .")
                        Call HideAllPanels()
                        lnkbtnCloseWindow.Visible = False
                        pnlCloseWindow.Visible = True
                    Else
                        WebMsgBox.Show("Sorry, that response was incorrect")
                    End If
                Else
                    WebMsgBox.Show("Please enter your response")
                End If
            End If
        End If
    End Sub
    
    Protected Function sBuildLogonDetailsMessageBody(ByVal dictAccounts As System.Collections.Generic.Dictionary(Of String, String)) As String
        Dim sbBody As New StringBuilder
        Dim sNewLine As String = "<br />" & Environment.NewLine
        sbBody.Append("Requested on:            ")
        sbBody.Append(DateTime.Now.ToString("F"))
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("Here are your logon details:")
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        For Each kvp As System.Collections.Generic.KeyValuePair(Of String, String) In dictAccounts
            sbBody.Append("UserID:                   ")
            sbBody.Append(kvp.Key)
            sbBody.Append(sNewLine)
            sbBody.Append("Password:                 ")
            sbBody.Append(kvp.Value)
            sbBody.Append(sNewLine)
            sbBody.Append(sNewLine)
        Next
        sBuildLogonDetailsMessageBody = sbBody.ToString
    End Function
    
    Protected Sub lbAccounts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As ListBox = sender
        pnUserKey = lb.SelectedValue
        Call HideAllPanels()
        pnlRequestLogonDetails.Visible = True
        Call SubmitRecoveryRequest(lb.SelectedValue)
    End Sub
    
    Property pdictAccounts() As Dictionary(Of String, String)
        Get
            Dim o As Object = ViewState("RC_Accounts")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CType(o, Dictionary(Of String, String))
        End Get
        Set(ByVal Value As Dictionary(Of String, String))
            ViewState("RC_Accounts") = Value
        End Set
    End Property
    
    Property pnUserKey() As Integer
        Get
            Dim o As Object = ViewState("RP_UserKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("RP_UserKey") = Value
        End Set
    End Property
    
    Property palMemorable() As ArrayList
        Get
            Dim o As Object = ViewState("RP_Memorable")
            If o Is Nothing Then
                Return Nothing
            End If
            Return o
        End Get
        Set(ByVal Value As ArrayList)
            ViewState("RP_Memorable") = Value
        End Set
    End Property

    Property psExpectedMemorableAnswer() As String
        Get
            Dim o As Object = ViewState("RP_ExpectedMemorableAnswer")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("RP_ExpectedMemorableAnswer") = Value
        End Set
    End Property

    Property psCurrentPassword() As String
        Get
            Dim o As Object = ViewState("RP_CurrentPassword")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("RP_CurrentPassword") = Value
        End Set
    End Property

    Property psCurrentUserId() As String
        Get
            Dim o As Object = ViewState("RP_CurrentUserId")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("RP_CurrentUserId") = Value
        End Set
    End Property

    Property pbLoggedOnAtSessionStart() As Boolean
        Get
            Dim o As Object = ViewState("RU_LoggedOn")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("RU_LoggedOn") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Access Request</title>
    <link href="tabs.css" rel="STYLESHEET" type="text/css" />
</head>
<body>
    <form id="frmUserIdApplication" runat="server">
        <asp:Panel id="pnlHeader" runat="server" visible="True" Width="100%">
            <table style="width: 100%; font-family:Verdana">
                <tr>
                    <td style="width: 25%">
                    </td>
                    <td style="width: 50%" align="center">
                        <asp:Image ID="imgLogo" runat="server" />
                    </td>
                    <td style="width: 25%">
                    </td>
                </tr>
                <tr>
                    <td colspan="3" align="center">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td colspan="3" align="right">
                        <br />
                        <asp:LinkButton ID="lnkbtnCloseWindow" runat="server" OnClientClick="window.close()" Font-Size="XX-Small" CausesValidation="False">close window</asp:LinkButton></td>
                </tr>
            </table>
        </asp:Panel>
        
        <asp:Panel id="pnlRequestLogonDetails" runat="server" visible="True" Width="100%">
            <table style="width: 100%; font-family:Verdana">
                <tr>
                    <td style="width: 5%">
                    </td>
                    <td style="width: 25%">
                    </td>
                    <td style="width: 70%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="lblLegendRequestLogonDetails" runat="server" Font-Size="X-Small" Text="Request login credentials" Font-Bold="True"></asp:Label></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="Label17" runat="server" Font-Size="XX-Small">Forgotten your User ID or password? Enter your email address then correctly answer the question that is displayed. Your login information will then be emailed to you.</asp:Label></td>
                </tr>
                <tr>
                    <td colspan="3">
                        <br />
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td valign="top" align="right">
                        <asp:Label ID="Label13" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                        <asp:Label ID="Label20" runat="server" Font-Size="XX-Small" ForeColor="Navy">Email Address</asp:Label></td>
                    <td>
                        <asp:TextBox runat="server" ID="tbRecoveryEmailAddr" Font-Size="XX-Small" ValidationGroup="vg1" Width="200px"></asp:TextBox>&nbsp;
                        <asp:RegularExpressionValidator ID="revRecoveryEmailAddr" runat="server" Font-Size="XX-Small" ControlToValidate="tbRecoveryEmailAddr"
                            ErrorMessage="Not a valid email address!" ValidationExpression="^([a-zA-Z0-9_'+*$%\^&!\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9:]{2,4})+$" ValidationGroup="vg1"></asp:RegularExpressionValidator>
                        <asp:RequiredFieldValidator ID="rfvEmailAddress" runat="server" Font-Size="XX-Small" ControlToValidate="tbRecoveryEmailAddr"
                            ValidationGroup="vg1"> Required field!</asp:RequiredFieldValidator></td>
                </tr>
                <tr id="trMemorableQuestion" visible="false" runat="server">
                    <td>
                    </td>
                    <td align="right" valign="top">
                        &nbsp;<asp:Label ID="lblMemorableQuestion" runat="server" Font-Size="XX-Small" ForeColor="Navy"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbMemorableAnswer" runat="server" Font-Size="XX-Small" ValidationGroup="vg1" Width="200px"/>
                        <asp:RequiredFieldValidator ID="rfvMemorableAnswer" runat="server" ControlToValidate="tbMemorableAnswer"
                            Font-Size="XX-Small" ValidationGroup="vg1"> Required field!</asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnSubmitRecoveryRequest" runat="server" Text="submit" ValidationGroup="vg1" OnClick="btnSubmitRecoveryRequest_Click" />
                        &nbsp;
                        <asp:Button ID="btnClearRecoveryRequest" runat="server" Text="clear" CausesValidation="False" /></td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;
                    </td>
                </tr>
                <tr>
                    <td colspan="3"><hr />
                    </td>
                </tr>
            </table>
        </asp:Panel>
        &nbsp;&nbsp;&nbsp;
        <asp:Panel ID="pnlSelectUserId" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
            <table style="width: 100%; font-family:Verdana">
                <tr>
                    <td style="width: 25%">
                        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Select account:"></asp:Label></td>
                    <td style="width: 50%" align="center">
                    </td>
                    <td style="width: 25%">
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label2" runat="server" Font-Names="Arial" Font-Size="XX-Small" Text="Two or more accounts have the same email address. Please select the account for which you require the credentials. The security question will come from this account."
                            Width="100%"></asp:Label></td>
                    <td>
                        <asp:ListBox ID="lbAccounts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoPostBack="True" OnSelectedIndexChanged="lbAccounts_SelectedIndexChanged"/>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    <asp:Panel ID="pnlCloseWindow" runat="server" Width="100%" Visible="false">
    <p>&nbsp;</p>    <p>&nbsp;</p>    <p>&nbsp;</p>    <p>&nbsp;</p>    <p>&nbsp;</p>
            <table style="width: 100%; font-family:Arial; font-size:xx-small">
                <tr>
                    <td style="width: 25%">
                    </td>
                    <td style="width: 50%" align="center">
                        <asp:Button ID="btnCloseWindow" OnClientClick="window.close()" runat="server" Text="close window" />
                    </td>
                    <td style="width: 25%">
                    </td>
                </tr>
            </table>
    </asp:Panel>
        <script type="text/javascript">
            document.frmUserIdApplication.txtFirstName.focus();
        </script>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
