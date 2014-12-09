<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Reference Control="LogonPanel.ascx" %>
<%@ import Namespace="SprintInternational" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<script runat="server">

    ' TO DO
    ' add check on UserId when creating default user
    ' test creating default user

    ' ExecuteNonQuery
    ' ExecuteQueryToListItemCollection
    ' ExecuteQueryToDataTable
    ' Testing comments for GitHub again

    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8
    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gnSiteKey As Integer
    Private gnMinPasswordLength As Integer
    Private gnMinPasswordUpperCaseChars As Integer
    Private gnMinPasswordLowerCaseChars As Integer
    Private gnMinPasswordDigits As Integer
    Private WithEvents Logon1 As ASP.LogonPanel
    
    Protected Sub Page_Load()
        If IsNothing(Session("SiteKey")) Then              ' is nothing on first load as panel has not yet been loaded to init SiteKey
            Call GetSiteKeyFromSiteName(sGetPath)
        End If
        gnSiteKey = Session("SiteKey")
        'Call LogInfo("Default.aspx ~ PageLoad ~ SiteKey = " & Session("SiteKey"))
        Call VerifyExistsPageContent()
        Call GetPageContent()
        Page.Header.Title = Session("SiteTitle") & " - Log in"
    End Sub

    'Protected Sub LogInfo(sInfo As String)
    '    Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('" & sInfo.Replace("'", "''") & "')")
    'End Sub
    
    Protected Function GetSiteKeyFromSiteName(ByVal sSiteName As String) As Integer
        'Call LogInfo("Default.aspx ~ GetSiteKeyFromSiteName ~ sSiteName = " & sSiteName)
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
                    Call RemoveLoginCookie()
                    Call ShowSiteSetup()
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
    
    Protected Sub RemoveLoginCookie()
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
    
    Protected Sub ShowSiteSetup()
        Call HideAllPanels()
        Call GetCustomerAccountCodes()
        pnlSiteSetup.Visible = True
        tbSiteName.Text = sGetPath()
    End Sub
    
    '    Protected Function sGetPath() As String
    '        Dim sPathInfo As String = Request.Path
    '        sGetPath = String.Empty
    '        If sPathInfo <> String.Empty Then
    '            sPathInfo = sPathInfo.Substring(1)
    '            Dim sPos As Integer = sPathInfo.IndexOf("/")
    '            If sPos > 0 Then
    '                sGetPath = sPathInfo.Substring(0, sPos)
    '            End If
    '        End If
    '    End Function

    Protected Function sGetPath() As String
        'Call LogInfo("Default.aspx ~ sGetPath ~ Request.Url.ToString.ToLower = " & Request.Url.ToString.ToLower)
        If Request.Url.ToString.ToLower.Contains("jupitermarketing") Then
            sGetPath = "jupiter"
        Else
            Dim sPathInfo As String = Request.Path
            sGetPath = String.Empty
            If sPathInfo <> String.Empty Then
                sPathInfo = sPathInfo.Substring(1)
                Dim sPos As Integer = sPathInfo.IndexOf("/")
                If sPos > 0 Then
                    sGetPath = sPathInfo.Substring(0, sPos)
                End If
            End If
        End If
    End Function
    
    Protected Sub VerifyExistsPageContent()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_SiteContent5", oConn)
        Try                                                 ' check an entry exists for this customer, create one if not
            oCmd.CommandType = CommandType.StoredProcedure

            oCmd.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Action").Value = "VERIFY"
                
            oCmd.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
            oCmd.Parameters("@SiteKey").Value = gnSiteKey
                
            oCmd.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@ContentType").Value = "VERIFY"
            
            oConn.Open()
            oCmd.ExecuteNonQuery()
            'Session.Remove("UserContentID")                 ' prevent write on first postback
        Catch ex As SqlException
            WebMsgBox.Show("VerifyExistsPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetPageContent()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent5", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = gnSiteKey
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "LPContent"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        Dim sTemp As String
        Logon1 = CType(LoadControl("~\LogonPanel.ascx"), ASP.LogonPanel)
        sTemp = dr("LgnPnlPane")
        Select Case CInt(sTemp)
            Case 0
                div_top.Controls.Add(Logon1)
            Case 1
                div_left_left.Controls.Add(Logon1)
            Case 2
                div_left_right.Controls.Add(Logon1)
            Case 3
                div_right_left.Controls.Add(Logon1)
            Case 4
                div_right_right.Controls.Add(Logon1)
            Case 5
                div_bottom.Controls.Add(Logon1)
            Case Else
                div_right_left.Controls.Add(Logon1)
        End Select
        '        div_all.InnerHtml = dr("LPAllContent") & String.Empty
        div_top.InnerHtml = dr("LPTopContent") & String.Empty
        span_left_left.InnerHtml = dr("LP1Content") & String.Empty
        span_left_right.InnerHtml = dr("LP2Content") & String.Empty
        span_right_left.InnerHtml = dr("LP3Content") & String.Empty
        span_right_right.InnerHtml = dr("LP4Content") & String.Empty
        div_bottom.InnerHtml = dr("LPBottomContent") & String.Empty
        Call ParseAttributes(dr("LPAllAttr") & String.Empty, div_all)
        Call ParseAttributes(dr("LPTopAttr") & String.Empty, div_top)
        Call ParseAttributes(dr("LPLeftAttr") & String.Empty, div_left)
        Call ParseAttributes(dr("LPRightAttr") & String.Empty, div_right)
        Call ParseAttributes(dr("LP1Attr") & String.Empty, div_left_left)
        Call ParseAttributes(dr("LP2Attr") & String.Empty, div_left_right)
        Call ParseAttributes(dr("LP3Attr") & String.Empty, div_right_left)
        Call ParseAttributes(dr("LP4Attr") & String.Empty, div_right_right)
        Call ParseAttributes(dr("LPBottomAttr") & String.Empty, div_bottom)
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
    
    Protected Sub HideAllPanels()
        pnlMain.Visible = False
        pnlMustChangePassword.Visible = False
        pnlSiteSetup.Visible = False
    End Sub

    Protected Sub HandleMustChangePassword() Handles Logon1.evMustChangePassword
        Call HideAllPanels()
        pnlMustChangePassword.Visible = True
        tbOldPassword.Focus()
    End Sub

    Protected Sub HandlePasswordexpired() Handles Logon1.evPasswordExpired
        Call HideAllPanels()
        pnlMustChangePassword.Visible = True
        lblPasswordExpired.Visible = True
        tbOldPassword.Focus()
    End Sub

    Protected Sub btnChangePassword_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbOldPassword.Focus()
        If tbNewPassword.Text <> tbConfirmNewPassword.Text Then
            lblPasswordChangeMessage.Text = "New password and confirmation do not match - password NOT changed!"
            Exit Sub
        End If
        If tbNewPassword.Text = tbOldPassword.Text Then
            lblPasswordChangeMessage.Text = "Your new password is the same as your old password - password NOT changed!"
            Exit Sub
        End If
        Call RetrieveCurrentLogonDetails()
        If tbNewPassword.Text = psCurrentPassword Then
            lblPasswordChangeMessage.Text = "Your new password is the same as your old password - password NOT changed!"
            Exit Sub
        End If
        If tbOldPassword.Text <> psCurrentPassword Then
            WebMsgBox.Show("The old password you entered did not match your current password - password NOT changed!")
            Exit Sub
        End If
        If tbNewPassword.Text = psCurrentUserId Then
            WebMsgBox.Show("Your password cannot be the same as your User ID - password NOT changed!")
            Exit Sub
        End If
        If bIsPasswordUsedPreviously(tbNewPassword.Text) Then
            WebMsgBox.Show("You have used this password previously - password NOT changed!")
            Exit Sub
        End If
        If tbNewPassword.Text.Length < gnMinPasswordLength Then
            WebMsgBox.Show("The new password you have specified is " & tbNewPassword.Text.Length & " character(s) long. The access policy in force specifies a minimum password length of " & gnMinPasswordLength & " characters. Password NOT changed!")
            Exit Sub
        End If
        If gnMinPasswordUpperCaseChars > 0 Then
            Dim i As Integer
            For Each c As Char In tbNewPassword.Text
                If c >= "A" And c <= "Z" Then
                    i += 1
                End If
            Next
            If i < gnMinPasswordUpperCaseChars Then
                Dim sPlural As String = String.Empty
                If gnMinPasswordUpperCaseChars > 1 Then
                    sPlural = "s"
                End If
                WebMsgBox.Show("The access policy in force requires a minimum of " & gnMinPasswordUpperCaseChars & " upper case character" & sPlural & ", which your new password does not supply. Password NOT changed!")
                Exit Sub
            End If
        End If
        If gnMinPasswordLowerCaseChars > 0 Then
            Dim i As Integer
            For Each c As Char In tbNewPassword.Text
                If c >= "a" And c <= "z" Then
                    i += 1
                End If
            Next
            If i < gnMinPasswordLowerCaseChars Then
                Dim sPlural As String = String.Empty
                If gnMinPasswordLowerCaseChars > 1 Then
                    sPlural = "s"
                End If
                WebMsgBox.Show("The access policy in force requires a minimum of " & gnMinPasswordLowerCaseChars & " lower case character" & sPlural & ", which your new password does not supply. Password NOT changed!")
                Exit Sub
            End If
        End If
        If gnMinPasswordDigits > 0 Then
            Dim i As Integer
            For Each c As Char In tbNewPassword.Text
                If c >= "0" And c <= "9" Then
                    i += 1
                End If
            Next
            If i < gnMinPasswordDigits Then
                Dim sPlural As String = String.Empty
                If gnMinPasswordDigits > 1 Then
                    sPlural = "s"
                End If
                WebMsgBox.Show("The access policy in force requires a minimum of " & gnMinPasswordDigits & " digit" & sPlural & ", which your new password does not supply. Password NOT changed!")
                Exit Sub
            End If
        End If
        Call SaveNewPassword(tbNewPassword.Text)
        Response.Redirect(ConfigLib.GetConfigItem_Default_Logon_Page)
    End Sub
    
    Protected Function bIsPasswordUsedPreviously(ByVal sPassword As String) As Boolean
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT * FROM PasswordHistory WHERE UserKey = " & Session("UserKey") & " AND Password = '" & sPassword & "'"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader
            bIsPasswordUsedPreviously = oDataReader.HasRows
        Catch ex As Exception
            WebMsgBox.Show("bIsPasswordUsedPreviously: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub lnkbtnResetPasswordFields_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbOldPassword.Text = String.Empty
        tbNewPassword.Text = String.Empty
        tbConfirmNewPassword.Text = String.Empty
    End Sub
    
    Protected Sub RetrieveCurrentLogonDetails()
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey2", oConn)
        
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
            If Not IsDBNull(oDataReader("MinPasswordLength")) Then
                gnMinPasswordLength = oDataReader("MinPasswordLength")
            End If
            If Not IsDBNull(oDataReader("MinPasswordUpperCaseChars")) Then
                gnMinPasswordUpperCaseChars = oDataReader("MinPasswordUpperCaseChars")
            End If
            If Not IsDBNull(oDataReader("MinPasswordLowerCaseChars")) Then
                gnMinPasswordLowerCaseChars = oDataReader("MinPasswordLowerCaseChars")
            End If
            If Not IsDBNull(oDataReader("MinPasswordDigits")) Then
                gnMinPasswordDigits = oDataReader("MinPasswordDigits")
            End If
        Catch ex As Exception
            WebMsgBox.Show("RetrieveCurrentLogonDetails: " & ex.Message)
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveNewPassword(ByVal sPlainPassword As String)
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        Dim sEncryptedPassword As String = oPassword.Encrypt(sPlainPassword)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SavePassword2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)
        
        Dim paramPlainPassword As SqlParameter = New SqlParameter("@PlainPassword", SqlDbType.NVarChar, 24)
        paramPlainPassword.Value = sPlainPassword
        oCmd.Parameters.Add(paramPlainPassword)
        
        Dim paramEncryptedPassword As SqlParameter = New SqlParameter("@EncryptedPassword", SqlDbType.NVarChar, 50)
        paramEncryptedPassword.Value = sEncryptedPassword
        oCmd.Parameters.Add(paramEncryptedPassword)
        
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SaveNewPassword: " & ex.Message)
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SendEmailAlert(ByVal sRecipient As String, ByVal sSubject As String, ByVal sText As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = "PAGE_ERROR_ALERT"
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oCmd.Parameters("@CustomerKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int, 4))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sText
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int, 4))
            oCmd.Parameters("@QueuedBy").Value = 0
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("SendEmailAlert: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub cbShowPasswords_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            tbOldPassword.TextMode = TextBoxMode.SingleLine
            tbNewPassword.TextMode = TextBoxMode.SingleLine
            tbConfirmNewPassword.TextMode = TextBoxMode.SingleLine
        Else
            tbOldPassword.TextMode = TextBoxMode.Password
            tbNewPassword.TextMode = TextBoxMode.Password
            tbConfirmNewPassword.TextMode = TextBoxMode.Password
        End If
        tbOldPassword.Focus()
    End Sub
    
    Property psCurrentPassword() As String
        Get
            Dim o As Object = ViewState("DF_CurrentPassword")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("DF_CurrentPassword") = Value
        End Set
    End Property

    Property psCurrentUserId() As String
        Get
            Dim o As Object = ViewState("DF_CurrentUserId")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("DF_CurrentUserId") = Value
        End Set
    End Property

    Property pnPasswordRetries() As Integer
        Get
            Dim o As Object = ViewState("DF_PasswordRetries")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("DF_PasswordRetries") = Value
        End Set
    End Property
    
    Protected Sub btnSaveSiteMapping_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TrimFields()
        If tbSecurityPassword.Text.Trim.ToLower = "tw140rn" Then
            If cbAddDefaultUser.Checked AndAlso ddlCustomerAccountCodes.SelectedIndex = 0 Then
                WebMsgBox.Show("Please select customer name for user account creation")
            Else
                If cbAddDefaultUser.Checked AndAlso bUserIdExists(tbUserId.Text) Then
                    WebMsgBox.Show("The UserId specified already exists - cannot continue")
                    Exit Sub
                End If
                If cbAddDefaultUser.Checked AndAlso bPermissionsOverlap(ddlCustomerAccountCodes.SelectedValue) Then
                    WebMsgBox.Show("Another user on this account already has one or more permissions that would be assigned to the default user - cannot continue")
                    Exit Sub
                End If
                Call SaveSiteConfiguration()
            End If
        Else
            WebMsgBox.Show("Could not save configuration changes")
            End If
    End Sub
    
    Protected Function bPermissionsOverlap(ByVal sCustomerKey As String) As Boolean
        bPermissionsOverlap = False
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT UserPermissions FROM UserProfile WHERE Type = 'SuperUser' AND CustomerKey = " & sCustomerKey)
        For Each dr As DataRow In oDataTable.Rows
            Dim nUserPermissions As Integer = dr("UserPermissions")
            If nUserPermissions And USER_PERMISSION_SITE_EDITOR Then
                bPermissionsOverlap = True
            End If
            If nUserPermissions And USER_PERMISSION_SITE_ADMINISTRATOR Then
                bPermissionsOverlap = True
            End If
            If nUserPermissions And USER_PERMISSION_ACCOUNT_HANDLER Then
                bPermissionsOverlap = True
            End If
        Next
    End Function

    Protected Sub TrimFields()
        tbFirstName.Text = tbFirstName.Text.Trim
        tbLastName.Text = tbLastName.Text.Trim
        tbUserId.Text = tbUserId.Text.Trim
        tbPassword.Text = tbPassword.Text.Trim
        tbEmailAddr.Text = tbEmailAddr.Text.Trim
    End Sub
    
    Protected Function bUserIdExists(ByVal sUserId As String) As Boolean
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT UserId FROM UserProfile WHERE UserId = '" & sUserId & "'"
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader
            bUserIdExists = oDataReader.HasRows
        Catch ex As Exception
            WebMsgBox.Show("bUserExists: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SaveSiteConfiguration()
        Call ExecuteNonQuery("INSERT INTO SiteNameToKeyMap (Path, SiteKey) VALUES ('" & tbSiteName.Text & "', " & CInt(tbSiteKey.Text) & ")")
        'Dim oConn As New SqlConnection(gsConn)
        'Dim sSQL As String = "INSERT INTO SiteNameToKeyMap (Path, SiteKey) VALUES ('" & tbSiteName.Text & "', " & CInt(tbSiteKey.Text) & ")"
        'Dim oCmd As New SqlCommand(sSQL, oConn)
        'Try
        '    oConn.Open()
        '    oCmd.ExecuteNonQuery()
        'Catch ex As SqlException
        '    WebMsgBox.Show("Error in SaveSiteConfiguration: " & ex.Message)
        'Finally
        '    oConn.Close()
        'End Try
        
        Session("SiteKey") = CInt(tbSiteKey.Text)
        
        Dim sXMLRotatorConfigSourceFilePath As String
        Dim sXMLNewsContentSourceFilePath As String
        Dim sXMLRotatorConfigDestinationFilePath As String
        Dim sXMLNewsContentDestinationFilePath As String
        sXMLRotatorConfigSourceFilePath = ".\rotator\news_config0" & ".xml"
        sXMLNewsContentSourceFilePath = ".\rotator\news0" & ".xml"
        sXMLRotatorConfigDestinationFilePath = ".\rotator\news_config" & Session("SiteKey") & ".xml"
        sXMLNewsContentDestinationFilePath = ".\rotator\news" & Session("SiteKey") & ".xml"
        If Not My.Computer.FileSystem.FileExists(MapPath(sXMLRotatorConfigDestinationFilePath)) Then
            My.Computer.FileSystem.CopyFile(MapPath(sXMLRotatorConfigSourceFilePath), MapPath(sXMLRotatorConfigDestinationFilePath))
        End If
        If Not My.Computer.FileSystem.FileExists(MapPath(sXMLNewsContentDestinationFilePath)) Then
            My.Computer.FileSystem.CopyFile(MapPath(sXMLNewsContentSourceFilePath), MapPath(sXMLNewsContentDestinationFilePath))
        End If
        If cbAddDefaultUser.Checked Then
            Call AddNewUser()
        End If
        Server.Transfer("session_expired.aspx")
    End Sub
    
    Protected Sub btnContinueWithDefault_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("SiteKey") = 0
        Server.Transfer("session_expired.aspx")
    End Sub
    
    Protected Sub GetCustomerAccountCodes()
        Dim oConn As New SqlConnection(gsConn)
        ddlCustomerAccountCodes.Items.Clear()
        Dim oCmd As New SqlCommand("spASPNET_Customer_GetActiveCustomerCodes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            ddlCustomerAccountCodes.DataSource = oCmd.ExecuteReader()
            ddlCustomerAccountCodes.DataTextField = "CustomerAccountCode"
            ddlCustomerAccountCodes.DataValueField = "CustomerKey"
            ddlCustomerAccountCodes.DataBind()
            ddlCustomerAccountCodes.Items.Insert(0, New ListItem("- select a customer -", 0))
        Catch ex As Exception
            WebMsgBox.Show("GetCustomerAccountCodes: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ddlCustomerAccountCodes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            lblCustomerKey.Text = ddl.SelectedValue
            tbSiteKey.Text = ddl.SelectedValue
        Else
            lblCustomerKey.Text = String.Empty
        End If
    End Sub
    
    Protected Sub AddNewUser()
        If Page.IsValid Then
            Dim bError As Boolean
            tbUserId.Text = tbUserId.Text.Trim
            If tbUserId.Text.ToLower = "sa" Then
                WebMsgBox.Show("SA is a reserved User ID, please reselect")
                Exit Sub
            End If
            Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
            Dim oConn As New SqlConnection(gsConn)
            
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_Add3", oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = 0
            oCmd.Parameters.Add(paramUserKey)
            Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 20)
            paramUserId.Value = tbUserId.Text
            oCmd.Parameters.Add(paramUserId)
            Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
            paramPassword.Value = oPassword.Encrypt(tbPassword.Text)
            oCmd.Parameters.Add(paramPassword)
            Dim paramFirstName As SqlParameter = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
            paramFirstName.Value = tbFirstName.Text
            oCmd.Parameters.Add(paramFirstName)
            Dim paramLastName As SqlParameter = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
            paramLastName.Value = tbLastName.Text
            oCmd.Parameters.Add(paramLastName)
            Dim paramTitle As SqlParameter = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
            paramTitle.Value = Nothing
            oCmd.Parameters.Add(paramTitle)
            Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
            paramDepartment.Value = String.Empty
            oCmd.Parameters.Add(paramDepartment)
            Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
            paramStatus.Value = "Active"
            oCmd.Parameters.Add(paramStatus)
            Dim paramType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
            paramType.Value = "SuperUser"
            oCmd.Parameters.Add(paramType)
            Dim paramCustomer As SqlParameter = New SqlParameter("@Customer", SqlDbType.Bit)
            paramCustomer.Value = 0
            oCmd.Parameters.Add(paramCustomer)
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            paramCustomerKey.Value = ddlCustomerAccountCodes.SelectedValue
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
            paramEmailAddr.Value = tbEmailAddr.Text
            oCmd.Parameters.Add(paramEmailAddr)
            Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
            paramTelephone.Value = String.Empty
            oCmd.Parameters.Add(paramTelephone)
            Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
            paramCollectionPoint.Value = String.Empty
            oCmd.Parameters.Add(paramCollectionPoint)
            Dim paramURL As SqlParameter = New SqlParameter("@URL", SqlDbType.NVarChar, 100)
            paramURL.Value = "default"
            oCmd.Parameters.Add(paramURL)

            Dim paramAbleToViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit)
            paramAbleToViewStock.Value = 0
            oCmd.Parameters.Add(paramAbleToViewStock)

            Dim paramAbleToCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit)
            paramAbleToCreateStockBooking.Value = 1
            oCmd.Parameters.Add(paramAbleToCreateStockBooking)

            Dim paramAbleToCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit)
            paramAbleToCreateCollectionRequest.Value = 1
            oCmd.Parameters.Add(paramAbleToCreateCollectionRequest)
            Dim paramAbleToViewGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit)
            paramAbleToViewGlobalAddressBook.Value = 1
            oCmd.Parameters.Add(paramAbleToViewGlobalAddressBook)
            Dim paramAbleToEditGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit)
            paramAbleToEditGlobalAddressBook.Value = 1
            oCmd.Parameters.Add(paramAbleToEditGlobalAddressBook)
            Dim paramRunningHeader As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
            paramRunningHeader.Value = "default"
            oCmd.Parameters.Add(paramRunningHeader)
            Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit)
            paramStockBookingAlert.Value = 1
            oCmd.Parameters.Add(paramStockBookingAlert)
            Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit)
            paramStockBookingAlertAll.Value = 1
            oCmd.Parameters.Add(paramStockBookingAlertAll)
            Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit)
            paramStockArrivalAlert.Value = 1
            oCmd.Parameters.Add(paramStockArrivalAlert)
            Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit)
            paramLowStockAlert.Value = 1
            oCmd.Parameters.Add(paramLowStockAlert)
            Dim paramCourierBookingAlert As SqlParameter = New SqlParameter("@ConsignmentBookingAlert", SqlDbType.Bit)
            paramCourierBookingAlert.Value = 1
            oCmd.Parameters.Add(paramCourierBookingAlert)
            Dim paramCourierBookingAlertAll As SqlParameter = New SqlParameter("@ConsignmentBookingAlertAll", SqlDbType.Bit)
            paramCourierBookingAlertAll.Value = 1
            oCmd.Parameters.Add(paramCourierBookingAlertAll)
            Dim paramCourierDespatchAlert As SqlParameter = New SqlParameter("@ConsignmentDespatchAlert", SqlDbType.Bit)
            paramCourierDespatchAlert.Value = 1
            oCmd.Parameters.Add(paramCourierDespatchAlert)
            Dim paramCourierDeliveryAlert As SqlParameter = New SqlParameter("@ConsignmentDeliveryAlert", SqlDbType.Bit)
            paramCourierDeliveryAlert.Value = 1
            oCmd.Parameters.Add(paramCourierDeliveryAlert)
            Dim paramUserPermissions As SqlParameter = New SqlParameter("@UserPermissions", SqlDbType.Int)
            paramUserPermissions.Value = 11        ' USER_PERMISSION_ACCOUNT_HANDLER (1) + USER_PERMISSION_SITE_ADMINISTRATOR (2) + USER_PERMISSION_SITE_EDITOR (8)
            oCmd.Parameters.Add(paramUserPermissions)
            Dim paramKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
            paramKey.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                oCmd.Connection = oConn
                oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                oTrans.Commit()
            Catch ex As SqlException
                bError = True
                oTrans.Rollback("AddRecord")
                If ex.Number = 2627 Then
                    WebMsgBox.Show("This User ID is already taken. Please select another User ID")
                    Exit Sub
                Else
                    WebMsgBox.Show("AddNewUser: " & ex.ToString)
                End If
            Finally
                oConn.Close()
            End Try
            If Not bError Then
                WebMsgBox.Show("User details added")
            End If
        End If
    End Sub
    
    Protected Sub cbAddDefaultUser_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            rfvFirstName.Enabled = True
            rfvLastName.Enabled = True
            rfvUserId.Enabled = True
            rfvPassword.Enabled = True
            rfvEmailAddr.Enabled = True
        Else
            rfvFirstName.Enabled = False
            rfvLastName.Enabled = False
            rfvUserId.Enabled = False
            rfvPassword.Enabled = False
            rfvEmailAddr.Enabled = False
        End If
    End Sub
    
    Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
        ExecuteNonQuery = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            oCmd.ExecuteNonQuery()
            ExecuteNonQuery = True
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
    Protected Function ExecuteQueryToListItemCollection(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If Not IsDBNull(oDataReader(sTextFieldName)) Then
                        sTextField = oDataReader(sTextFieldName)
                    Else
                        sTextField = String.Empty
                    End If
                    If Not IsDBNull(oDataReader(sValueFieldName)) Then
                        sValueField = oDataReader(sValueFieldName)
                    Else
                        sValueField = String.Empty
                    End If
                    oListItemCollection.Add(New ListItem(sTextField, sValueField))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

    Protected Function ExecuteQueryToDataTable(ByVal sQuery As String) As DataTable
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
        ExecuteQueryToDataTable = oDataTable
    End Function

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Log In</title>
</head>
<body>
    <form id="frmlogon" runat="server">
        <asp:Panel ID="pnlMain" runat="server" Visible="true" Width="100%">
            <div id="div_all" runat="server">
                <div id="div_top" runat="server">
                    <asp:PlaceHolder ID="ph_top" runat="server" />
                </div>
                <div id="div_left" runat="server">
                    <div id="div_left_left" runat="server">
                        <span id="span_left_left" runat="server">
                            <asp:PlaceHolder ID="ph_left_left" runat="server" />
                        </span>
                    </div>
                    <div id="div_left_right" runat="server">    
                        <span id="span_left_right" runat="server">
                            <asp:PlaceHolder ID="ph_left_right" runat="server" />
                        </span>
                    </div>
                </div>
                <div id="div_right" runat="server">
                    <div id="div_right_left" runat="server">
                        <span id="span_right_left" runat="server">
                            <asp:PlaceHolder ID="ph_right_left" runat="server" />
                        </span>
                    </div>
                    <div id="div_right_right" runat="server">
                        <span id="span_right_right" runat="server">
                            <asp:PlaceHolder ID="ph_right_right" runat="server" />
                        </span>
                    </div>
                </div>
                <div id="div_bottom" runat="server">
                    <asp:PlaceHolder ID="ph_bottom" runat="server" />
                </div>
            </div>
        </asp:Panel>
        <asp:Panel ID="pnlMustChangePassword" Font-Names="Verdana" Font-Size="XX-Small" Visible="false" runat="server" Width="100%">
            <strong>
                <asp:Label ID="lblPasswordExpired" runat="server" Font-Bold="True" Font-Names="Verdana"
                    Font-Size="XX-Small" Visible="False">Your password has expired. </asp:Label>You must set a new password for this account before you can continue using the system.
                When you have set a new password you must log in again using your new password.</strong><br />
            <br />
            <strong>Change Password</strong><br />
            <table style="width: 100%; font-family:Verdana; font-size:xx-small">
                <tr>
                    <td style="width: 5%">&nbsp;
                    </td>
                    <td style="width: 20%; white-space:nowrap">
                    </td>
                    <td style="width: 75%">
                        <asp:CheckBox ID="cbShowPasswords" runat="server" AutoPostBack="True" OnCheckedChanged="cbShowPasswords_CheckedChanged"
                            Text="show passwords" /></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        Old Password:
                    </td>
                    <td>
                        <asp:TextBox ID="tbOldPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="150px" TextMode="Password" MaxLength="12"></asp:TextBox>&nbsp;
                        <asp:RequiredFieldValidator ID="rfvOldPassword" runat="server" ControlToValidate="tbOldPassword"
                            ErrorMessage="Required field!" Font-Size="XX-Small" ValidationGroup="vg2"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        New Password:</td>
                    <td>
                        <asp:TextBox ID="tbNewPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="150px" TextMode="Password" MaxLength="12"></asp:TextBox>&nbsp;
                        <asp:RequiredFieldValidator ID="rfvNewPassword" runat="server" ControlToValidate="tbNewPassword"
                            ErrorMessage="Required field!" Font-Size="XX-Small" ValidationGroup="vg2"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        Confirm New Password:</td>
                    <td>
                        <asp:TextBox ID="tbConfirmNewPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="150px" TextMode="Password" MaxLength="12"></asp:TextBox>&nbsp;
                        <asp:RequiredFieldValidator ID="rfvConfirmNewPassword" runat="server" ControlToValidate="tbConfirmNewPassword"
                            ErrorMessage="Required field!" Font-Size="XX-Small" ValidationGroup="vg2"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnChangePassword" runat="server" Text="change password" OnClick="btnChangePassword_Click" />
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        <asp:LinkButton ID="lnkbtnResetPasswordFields" runat="server" OnClick="lnkbtnResetPasswordFields_Click">reset password fields</asp:LinkButton></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="lblPasswordChangeMessage" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlSiteSetup" Visible="false" runat="server" Width="100%">
            <asp:Label ID="Label1" runat="server" Text="Define Site Mapping & Default User" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label><br />
            <br />
            <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="This site does not have a mapping from site name to site key. You can define the mapping here."></asp:Label><br />
            <br />
            <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Normally the site key is the same as the customer key, so that the customer's URL goes to a unique site, but you can point two or more different URLs to the same site by assigning them the same site key."></asp:Label><br />
            <br />
            <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="You can also define a new user account to be the Account Handler, Site Administrator and Site Editor."></asp:Label><br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 30%">
                    </td>
                    <td style="width: 70%">
                    </td>
                </tr>
                <tr>
                    <td align="right" style="height: 21px">
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Find customer key from account code:"></asp:Label>
                    </td>
                    <td style="height: 21px">
                        <asp:DropDownList ID="ddlCustomerAccountCodes" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCustomerAccountCodes_SelectedIndexChanged">
                        </asp:DropDownList>
                        <asp:Label ID="lblCustomerKey" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Site Name:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbSiteName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvSiteName" runat="server" ControlToValidate="tbSiteName"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Site Key:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbSiteKey" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="tbSiteKey"
                            ErrorMessage="must be a number between 0 & 9999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="9999" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="rfvSiteKey" runat="server" ControlToValidate="tbSiteKey"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbAddDefaultUser" runat="server" Checked="True" Font-Names="Verdana"
                            Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbAddDefaultUser_CheckedChanged" Text="Add default user" />
                        &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                    </td>
                    <td>
                        <asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="New internal user will be created for the account selected above,<br />and given Account Handler, Site Administrator & Site Editor privileges."></asp:Label><br />
                        <asp:Label ID="Label15" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            ForeColor="Red" Text="THIS IS A PRIVILEGED ACCOUNT WHICH MUST NOT BE GIVEN TO CUSTOMERS !!!"></asp:Label></td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="First Name:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbFirstName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="250px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvFirstName" runat="server" ControlToValidate="tbFirstName"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Last Name:"></asp:Label></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbLastName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="250px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvLastName" runat="server" ControlToValidate="tbLastName"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="User Id:"></asp:Label></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbUserId" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="250px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvUserId" runat="server" ControlToValidate="tbUserId"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Password:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="250px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvPassword" runat="server" ControlToValidate="tbPassword"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Email Addr:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbEmailAddr" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="250px">account.managers@transworld.eu.com</asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvEmailAddr" runat="server" ControlToValidate="tbEmailAddr"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                    </td>
                    <td>
                    &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Router access:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbSecurityPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvSecurityPassword" runat="server" ControlToValidate="tbSecurityPassword"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Button ID="btnSaveSiteMapping" runat="server" OnClick="btnSaveSiteMapping_Click"
                            Text="save configuration" /></td>
                    <td>
                        &nbsp;<asp:Button ID="btnContinueWithDefault" runat="server" CausesValidation="False" OnClientClick='return confirm("Hmmm, are you ABSOLUTELY sure you want to continue with site key = 0?");'
                            OnClick="btnContinueWithDefault_Click" Text="continue with default (site key = 0)" /></td>
                </tr>
            </table>
        </asp:Panel>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
