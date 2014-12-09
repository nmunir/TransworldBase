<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Const TRANSWORLD_RECIPIENT As String = "account.managers@transworld.eu.com"
    Const KODAK_DFIS_RECIPIENT As String = "andreana.scott@kodak.com"
    Const RAMBLERS_RECIPIENT As String = "adrianne.thyer@transworld.eu.com"

    Const HYSTERYALE_RECIPIENT As String = "account.managers@transworld.eu.com"
    Const HYSTERYALE_RECIPIENT_CLIENT As String = " luis.garcia@nmhg.com"

    Const SENTFROM_EMAIL_ADDR As String = "automailer@transworld.eu.com"

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            imgLogo.ImageUrl = "http://my.transworld.eu.com/common/images/logos/transworld.jpg"
            Call HideAllPanels()
            If Request.QueryString("type") = "newuser" Then
                pnlRequestAccess.Visible = True
                pnlRequestLogonDetails.Visible = False
                pnlPrivacyStatement.Visible = True
                If IsHysterOrYale() Then
                    trDepartment.Visible = False
                    trRegionalOffice.Visible = False
                    trDealerShipCompanyName.Visible = True
                    trDealershipCodeNACCOLocationCode.Visible = True
                    trNACCODepartmentCode.Visible = True
                End If
                If IsRamblers() Then
                    trJobTitle.Visible = False
                    trDepartment.Visible = False
                    trRegionalOffice.Visible = False
                    trRamblersAreaGroup.Visible = True
                End If
                If IsWU() Then
                    trDepartment.Visible = False
                    trRegionalOffice.Visible = False
                    trJobTitle.Visible = False
                    trTerminalID.Visible = True
                End If
            Else
                pnlRequestLogonDetails.Visible = True
                pnlPrivacyStatement.Visible = True
                tbRecoveryEmailAddr.Focus()
            End If
        End If
        tbRecoveryEmailAddr.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitRecoveryRequest.ClientID + "')")
        txtFirstName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtLastName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtJobTitle.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtDepartment.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtRegionalOffice.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtEmail.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtComments.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSubmitAccessRequest.ClientID + "')")
        txtFirstName.Focus()
    End Sub
    
    Protected Sub HideAllPanels()
        pnlRequestAccess.Visible = False
        pnlRequestLogonDetails.Visible = False
        pnlCloseWindow.Visible = False
        pnlPrivacyStatement.Visible = False
    End Sub
    
    Protected Function IsHysterOrYale() As Boolean
        IsHysterOrYale = sGetPath.ToLower.Contains("hyster") Or sGetPath.ToLower.Contains("yale")
    End Function
    
    Protected Function IsRamblers() As Boolean
        IsRamblers = sGetPath.ToLower.Contains("ramblers")
    End Function
    
    Protected Function IsWU() As Boolean
        IsWU = sGetPath.ToLower.Contains("wurs") Or sGetPath.ToLower.Contains("wuire")
    End Function
    
    Protected Sub btnResetAccessRequest_Click(ByVal sender As Object, ByVal e As EventArgs)
        txtFirstName.Text = ""
        txtLastName.Text = ""
        txtJobTitle.Text = ""
        txtDepartment.Text = ""
        txtRegionalOffice.Text = ""
        txtEmail.Text = ""
        txtComments.Text = ""
        tbDealershipCompanyName.Text = ""
        tbDealershipCodeNACCOLocationCode.Text = ""
        tbNACCODepartmentCode.Text = ""
        txtFirstName.Focus()
    End Sub

    Protected Function sGetPath() As String
        sGetPath = String.Empty
        Dim sPathInfo As String = Request.Path
        Dim sURL As String = Request.Url.ToString
        If sURL.Contains("kodakpos") Then
            sGetPath = "Kodak DFIS (kodakpos.co.uk)"
        ElseIf sURL.Contains("jupitermarketing") Then
            sGetPath = "Jupiter (jupitermarketing.co.uk)"
        Else
            If sPathInfo <> String.Empty Then
                sPathInfo = sPathInfo.Substring(1)
                Dim sPos As Integer = sPathInfo.IndexOf("/")
                If sPos > 0 Then
                    sGetPath = sPathInfo.Substring(0, sPos)
                End If
            End If
        End If
    End Function

    Protected Function sBuildStandardAccessRequestMessageBody() As String
        Dim sbBody As New StringBuilder
        Dim sNewLine As String = "<br />" & Environment.NewLine
        sbBody.Append("Requested on:            ")
        sbBody.Append(DateTime.Now.ToString("F"))
        sbBody.Append(sNewLine)
        sbBody.Append("Requesting site:         ")
        sbBody.Append(sGetPath.ToUpper)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("First name:              ")
        sbBody.Append(txtFirstName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Last name:               ")
        sbBody.Append(txtLastName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Job title:               ")
        sbBody.Append(txtJobTitle.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Dept/Cost Centre:        ")
        sbBody.Append(txtDepartment.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Regional office:         ")
        sbBody.Append(txtRegionalOffice.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Email Address:           ")
        sbBody.Append("mailto:" & txtEmail.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Comments:                ")
        sbBody.Append(txtComments.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("This message was automatically generated - please do not reply")
        sbBody.Append(sNewLine)
        sBuildStandardAccessRequestMessageBody = sbBody.ToString
    End Function
    
    Protected Function sBuildWUAccessRequestMessageBody() As String
        Dim sbBody As New StringBuilder
        Dim sNewLine As String = "<br />" & Environment.NewLine
        sbBody.Append("Requested on:            ")
        sbBody.Append(DateTime.Now.ToString("F"))
        sbBody.Append(sNewLine)
        sbBody.Append("Requesting site:         ")
        sbBody.Append(sGetPath.ToUpper)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("First name:              ")
        sbBody.Append(txtFirstName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Last name:               ")
        sbBody.Append(txtLastName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Terminal ID:             ")
        sbBody.Append(txtTerminalID.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Email Address:           ")
        sbBody.Append("mailto:" & txtEmail.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Comments:                ")
        sbBody.Append(txtComments.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("This message was automatically generated - please do not reply")
        sbBody.Append(sNewLine)
        sBuildWUAccessRequestMessageBody = sbBody.ToString
    End Function
    
    Protected Function sBuildHysterYaleAccessRequestMessageBody() As String
        Dim sbBody As New StringBuilder
        Dim sNewLine As String = "<br />" & Environment.NewLine
        sbBody.Append("Requested on:                          ")
        sbBody.Append(DateTime.Now.ToString("F"))
        sbBody.Append(sNewLine)
        sbBody.Append("Requesting site:                       ")
        sbBody.Append(sGetPath.ToUpper)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("First name:                            ")
        sbBody.Append(txtFirstName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Last name:                             ")
        sbBody.Append(txtLastName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Job title:                             ")
        sbBody.Append(txtJobTitle.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Dealership / Company Name:             ")
        sbBody.Append(tbDealershipCompanyName.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Dealership Code / NACCO Location Code: ")
        sbBody.Append(tbDealershipCodeNACCOLocationCode.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("NACCO Deparment Code (if relevant):    ")
        sbBody.Append(tbNACCODepartmentCode.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Email Address:                         ")
        sbBody.Append("mailto:" & txtEmail.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append("Comments:                              ")
        sbBody.Append(txtComments.Text.Trim)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append(sNewLine)
        sbBody.Append("This message was automatically generated - please do not reply")
        sbBody.Append(sNewLine)
        sBuildHysterYaleAccessRequestMessageBody = sbBody.ToString
    End Function
    
    Protected Sub btnSubmitAccessRequest_Click(ByVal sender As Object, ByVal e As EventArgs)
        If IsValid Then
            Dim sMessage As String
            If IsHysterOrYale() Then
                sMessage = sBuildHysterYaleAccessRequestMessageBody()
            ElseIf IsWU() Then
                sMessage = sBuildWUAccessRequestMessageBody()
            Else
                sMessage = sBuildStandardAccessRequestMessageBody()
            End If
            Dim sOriginator As String = sGetPath()
            If sOriginator.Contains("kodakpos") Then
                Call SendHTMLEmail("WEB_USERID_REQUEST", 0, 0, 0, 0, KODAK_DFIS_RECIPIENT, "Request for stock system account from www.kodakpos.co.uk", sMessage, sMessage, 0)
            ElseIf IsHysterOrYale() Then
                Call SendHTMLEmail("WEB_USERID_REQUEST", 0, 0, 0, 0, HYSTERYALE_RECIPIENT, "Request for stock system account from site " & sOriginator.ToUpper, sMessage, sMessage, 0)
                Call SendHTMLEmail("WEB_USERID_REQUEST", 0, 0, 0, 0, HYSTERYALE_RECIPIENT_CLIENT, "Request for stock system account from site " & sOriginator.ToUpper, sMessage, sMessage, 0)
            ElseIf IsRamblers() Then
                Call SendHTMLEmail("WEB_USERID_REQUEST", 0, 0, 0, 0, RAMBLERS_RECIPIENT, "Request for stock system account from site " & sOriginator.ToUpper, sMessage, sMessage, 0)
            Else
                Call SendHTMLEmail("WEB_USERID_REQUEST", 0, 0, 0, 0, TRANSWORLD_RECIPIENT, "Request for stock system account from site " & sOriginator.ToUpper, sMessage, sMessage, 0)
            End If
            WebMsgBox.Show("Thank you.  Your application will be processed shortly.")
            Call HideAllPanels()
            lnkbtnCloseWindow.Visible = False
            pnlCloseWindow.Visible = True
        End If
    End Sub
    
    Protected Sub SendHTMLEmail(ByVal sType As String, ByVal nCustomerKey As Integer, ByVal nStockBookingKey As Integer, ByVal nConsignmentKey As Integer, ByVal nProductKey As Integer, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String, ByVal nQueuedBy As Integer)
        Dim bError As Boolean = False

        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oCmd.Parameters("@CustomerKey").Value = nCustomerKey
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int, 4))
            oCmd.Parameters("@StockBookingKey").Value = nStockBookingKey
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ConsignmentKey").Value = nConsignmentKey
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ProductKey").Value = nProductKey
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int, 4))
            oCmd.Parameters("@QueuedBy").Value = nQueuedBy
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SendHTMLEmail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub RetrieveCurrentLogonDetails()
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey", oConn)
        
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = Session("UserKey")
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
        Catch ex As Exception
            WebMsgBox.Show("Unable to continue due to an internal error")
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveNewPassword(ByVal sPlainTextPassword As String)
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim sEncryptedPassword As String = oPassword.Encrypt(sPlainTextPassword)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SavePassword", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)
        Dim paramPassword As SqlParameter = New SqlParameter("@EncryptedPassword", SqlDbType.NVarChar, 24)
        paramPassword.Value = sEncryptedPassword
        oCmd.Parameters.Add(paramPassword)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Unable to continue due to an internal error")
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
                Loop
            Else
            End If
        Catch ex As Exception
            WebMsgBox.Show("Unable to continue due to an internal error")
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
        GetMatchingEmailAddresses = dictAccounts
    End Function

    Protected Sub btnSubmitRecoveryRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsValid Then
            Dim dictAccounts As System.Collections.Generic.Dictionary(Of String, String) = GetMatchingEmailAddresses()
            If dictAccounts.Count = 0 Then
                WebMsgBox.Show("No account matches this email address")
            Else
                Dim sMessageBody = sBuildLogonDetailsMessageBody(dictAccounts)
                Call SendHTMLEmail("WEB_LOGIN_REQUEST", 0, 0, 0, 0, tbRecoveryEmailAddr.Text, "Request for Transworld login credentials", sMessageBody, sMessageBody, 0)
                WebMsgBox.Show("An email containing your login details has been sent to " & tbRecoveryEmailAddr.Text & " .")
                Call HideAllPanels()
                lnkbtnCloseWindow.Visible = False
                pnlCloseWindow.Visible = True
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
    
    Protected Sub lnkbtnRamblersGroup_Click(sender As Object, e As System.EventArgs)
        Dim lb As LinkButton = sender
        xdsVar26RamblersAreaGroups.XPath = "RamblersAreaGroups/Ramblers" & lb.CommandArgument & "/areaGroup"
        xdsVar26RamblersAreaGroups.DataBind()
    End Sub
    
    Property psCurrentPassword() As String
        Get
            Dim o As Object = ViewState("RU_CurrentPassword")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("RU_CurrentPassword") = Value
        End Set
    End Property

    Property psCurrentUserId() As String
        Get
            Dim o As Object = ViewState("RU_CurrentUserId")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("RU_CurrentUserId") = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Access Request</title>
    <link href="sprint.css" rel="STYLESHEET" type="text/css" />
</head>
<body>
    <form id="frmUserIdApplication" runat="server">
    <asp:Panel ID="pnlHeader" runat="server" Visible="True" Width="100%">
        <table style="width: 100%; font-family: Verdana">
            <tr>
                <td style="width: 25%">
                </td>
                <td style="width: 50%" align="center">
                    <asp:Image ID="imgLogo" runat="server" 
                        ImageUrl="http://my.transworld.eu.com/common/images/logos/transworld.jpg" />
                </td>
                <td style="width: 25%">
                </td>
            </tr>
            <tr>
                <td colspan="3" align="center">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="3" align="right">
                    <br />
                    <asp:LinkButton ID="lnkbtnCloseWindow" runat="server" OnClientClick="window.close()"
                        Font-Size="XX-Small" CausesValidation="False">close window</asp:LinkButton>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlRequestLogonDetails" runat="server" Visible="True" Width="100%">
        <table style="width: 100%; font-family: Verdana">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 25%">
                </td>
                <td style="width: 172px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblLegendRequestLogonDetails" runat="server" Font-Size="X-Small" Text="REQUEST ACCOUNT LOGON DETAILS (existing users)"
                        Font-Bold="True"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="Label17" runat="server" Font-Size="XX-Small">Forgotten your User ID or password? Enter your email address below.  Your logon information will be emailed to you.</asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td valign="top">
                    <asp:Label ID="Label20" runat="server" Font-Size="XX-Small" ForeColor="Navy">Email Address</asp:Label>
                    <asp:Label ID="Label13" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                    &nbsp;&nbsp; &nbsp;&nbsp;
                </td>
                <td style="width: 172px">
                    <asp:TextBox runat="server" ID="tbRecoveryEmailAddr" Font-Size="XX-Small" ValidationGroup="vg1"
                        Font-Names="Verdana"></asp:TextBox>&nbsp;
                    <asp:RegularExpressionValidator ID="revRecoveryEmailAddr" runat="server" Font-Size="XX-Small"
                        ControlToValidate="tbRecoveryEmailAddr" ErrorMessage="Not a valid email address!"
                        ValidationExpression="^([a-zA-Z0-9_'+*$%\^&!\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9:]{2,4})+$"
                        ValidationGroup="vg1"></asp:RegularExpressionValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" Font-Size="XX-Small"
                        ControlToValidate="tbRecoveryEmailAddr" ValidationGroup="vg1"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td style="height: 26px">
                </td>
                <td style="height: 26px">
                </td>
                <td style="width: 172px; height: 26px;">
                    <asp:Button ID="btnSubmitRecoveryRequest" runat="server" Text="submit" ValidationGroup="vg1"
                        OnClick="btnSubmitRecoveryRequest_Click" />
                    &nbsp;
                    <asp:Button ID="btnClearRecoveryRequest" runat="server" Text="clear" CausesValidation="False" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <hr />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlRequestAccess" runat="server" Visible="False" Width="100%">
        <table style="width: 100%; font-family: Verdana">
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
                    <asp:Label ID="lblLegendRequestAccess" runat="server" Font-Size="X-Small" Text="REQUEST ACCESS (new users)"
                        Font-Bold="True"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="Label1" runat="server" Font-Size="XX-Small">To request access to the Transworld online ordering system, please complete the form below. Your User ID and password will be sent to you as soon as your application has been accepted.</asp:Label><br />
                </td>
            </tr>
            <tr>
                <td colspan="3" style="height: 38px">
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label3" runat="server" Font-Size="XX-Small" ForeColor="Navy">First Name</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label4" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ID="txtFirstName" Font-Size="XX-Small" ValidationGroup="vg3"
                        Width="200px" Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvFirstName" Font-Size="XX-Small" runat="server"
                        ControlToValidate="txtFirstName" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label5" runat="server" Font-Size="XX-Small" ForeColor="Navy">Last Name</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label6" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ID="txtLastName" Font-Size="XX-Small" ValidationGroup="vg3"
                        Width="200px" Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvLastName" runat="server" Font-Size="XX-Small"
                        ControlToValidate="txtLastName" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trTerminalID" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label70" runat="server" Font-Size="XX-Small" ForeColor="Navy">Terminal ID</asp:Label>
                    &nbsp;<asp:Label ID="Label71" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="txtTerminalID" Width="200px" Font-Names="Verdana"/>
                    <asp:RequiredFieldValidator ID="rfvTerminalID" runat="server" 
                        ControlToValidate="txtTerminalID" Font-Size="XX-Small" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trJobTitle" runat="server" visible="true">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label7" runat="server" Font-Size="XX-Small" ForeColor="Navy">Job Title</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="txtJobTitle" Width="200px" Font-Names="Verdana"></asp:TextBox>
                </td>
            </tr>
            <tr id="trRamblersAreaGroup" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label18" runat="server" Font-Size="XX-Small" ForeColor="Navy">Area / Group</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label19" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <br />
                    <asp:LinkButton ID="lnkbtnRamblersGroupAB" CommandArgument="A-B" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">A-B</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupCD" CommandArgument="C-D" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">C-D</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupEG" CommandArgument="E-G" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">E-G</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupHJ" CommandArgument="H-J" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">H-J</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupKL" CommandArgument="K-L" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">K-L</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupMN" CommandArgument="M-N" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">M-N</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupOR" CommandArgument="O-R" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">O-R</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupS" CommandArgument="S" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">S</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupTZ" CommandArgument="T-Z" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">T-Z</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupOther" CommandArgument="Other" runat="server"
                        OnClick="lnkbtnRamblersGroup_Click" Font-Names="Verdana" Font-Size="XX-Small"
                        CausesValidation="False">Other</asp:LinkButton>
                    <br />
                    <br />
                    <asp:DropDownList ID="ddlRamblersAreaGroup" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        DataSourceID="xdsVar26RamblersAreaGroups" DataTextField="name">
                        <asp:ListItem>- Select from the alphabetic choices above to find your area/group -</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;<asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration26RamblersAreaGroup"
                        runat="server" ControlToValidate="ddlRamblersAreaGroup" ErrorMessage="Required field!"
                        Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -"
                        ValidationGroup="vg3" Font-Bold="True" ForeColor="Red" />
                    <br />
                    <br />
                </td>
            </tr>
            <tr id="trDepartment" runat="server" visible="true">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label8" runat="server" Font-Size="XX-Small" ForeColor="Navy">Department / Cost Centre</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="txtDepartment" Width="200px"
                        Font-Names="Verdana"></asp:TextBox>
                </td>
            </tr>
            <tr id="trRegionalOffice" runat="server" visible="true">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label9" runat="server" Font-Size="XX-Small" ForeColor="Navy">Regional Office</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="txtRegionalOffice" Width="200px"
                        Font-Names="Verdana"></asp:TextBox>
                </td>
            </tr>
            <tr id="trDealerShipCompanyName" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label2" runat="server" Font-Size="XX-Small" ForeColor="Navy">Dealership / Company Name</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label21" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="tbDealershipCompanyName" Width="200px"
                        Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvDealershipCompanyName" runat="server" ControlToValidate="tbDealershipCompanyName"
                        Font-Size="XX-Small" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trDealershipCodeNACCOLocationCode" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label14" runat="server" Font-Size="XX-Small" ForeColor="Navy">Dealership Code / NACCO Location Code</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label22" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="tbDealershipCodeNACCOLocationCode"
                        Width="200px" Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvDealershipCodeNACCOLocationCode" runat="server"
                        ControlToValidate="tbDealershipCodeNACCOLocationCode" Font-Size="XX-Small" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr id="trNACCODepartmentCode" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label15" runat="server" Font-Size="XX-Small" ForeColor="Navy">NACCO Department Code (if relevant)</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" Font-Size="XX-Small" ID="tbNACCODepartmentCode" Width="200px"
                        Font-Names="Verdana"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label10" runat="server" Font-Size="XX-Small" ForeColor="Navy">Email Address</asp:Label>
                    &nbsp;
                    <asp:Label ID="Label11" runat="server" Font-Size="XX-Small" ForeColor="Red">*</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ID="txtEmail" Font-Size="XX-Small" ValidationGroup="vg3"
                        Width="200px" Font-Names="Verdana"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" Font-Size="XX-Small" runat="server"
                        ControlToValidate="txtEmail" ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"> Required field!</asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="revEmailAddr" runat="server" Font-Size="XX-Small"
                        ControlToValidate="txtEmail" ErrorMessage="Not a valid email address!" ValidationExpression="^([a-zA-Z0-9_'+*$%\^&!\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9:]{2,4})+$"
                        ValidationGroup="vg3" Font-Bold="True" ForeColor="Red"></asp:RegularExpressionValidator>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label16" runat="server" Font-Size="XX-Small" ForeColor="Navy">Comments</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" TextMode="MultiLine" Font-Size="XX-Small" Width="350px"
                        ID="txtComments" Rows="4" Font-Names="Verdana"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                    <br />
                    <asp:Button runat="server" ID="btnSubmitAccessRequest" Text="submit" OnClick="btnSubmitAccessRequest_Click"
                        ValidationGroup="vg3"></asp:Button>
                    &nbsp;
                    <asp:Button runat="server" ID="btnResetAccessRequest" CausesValidation="False" Text="clear"
                        OnClick="btnResetAccessRequest_Click"></asp:Button>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label runat="server" ID="lblRequestAccessMessage" ForeColor="Red" Font-Size="XX-Small"
                        Font-Names="Arial"></asp:Label>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlPrivacyStatement" runat="server" Width="100%" Visible="false">
        <br />
        <table style="width: 100%; font-family: Verdana; font-size: xx-small">
            <tr>
                <td style="width: 2%; height: 14px;">
                </td>
                <td style="width: 96%; height: 14px;">
                </td>
                <td style="width: 2%; height: 14px;">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="center">
                    <asp:Label ID="Label12" runat="server" Font-Size="XX-Small" ForeColor="Gray" Text="Privacy Statement: Your contact information will remain confidential and will be used by Transworld solely to create and maintain a profile for your own use."></asp:Label>
                </td>
                <td>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlCloseWindow" runat="server" Width="100%" Visible="false">
        <p>
            &nbsp;</p>
        <p>
            &nbsp;</p>
        <p>
            &nbsp;</p>
        <p>
            &nbsp;</p>
        <p>
            &nbsp;</p>
        <table style="width: 100%; font-family: Arial; font-size: xx-small">
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
    <asp:XmlDataSource ID="xdsVar26RamblersAreaGroups" runat="server" DataFile="~/on_line_picks_config_ramblers.xml"
        XPath="RamblersAreaGroups/RamblersA-B/areaGroup" />
    <script type="text/javascript">
        document.frmUserIdApplication.txtFirstName.focus();
    </script>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>