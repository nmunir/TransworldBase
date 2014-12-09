<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="SprintInternational" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ import Namespace="Microsoft.Win32" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>

<script runat="server">

    ' TO DO 18MAR11
    ' fix hidCustomer problem
    ' check ordering of new user entries added
    ' add dropdown to select number of user records displayed
    
    ' TO DO
    ' check if deleted customers are shown in dropdown & remove if they are
    ' Make site-configurable: Product Owners; Allow User to Edit Cost Centre; Cost Centre Label; # Sub Categories; 

    ' SPROCS
    ' spASPNET_UserProfile_GetProfileFromKey
    ' spASPNET_Customer_GetNameFromKey
    ' spASPNET_UserProfile_GetProfileFromKey
    ' spASPNET_Customer_GetUserProfiles
    ' spASPNET_Customer_GetActiveUserProfiles
    ' spASPNET_UserProfile_GetProductProfileFromKey
    ' spASPNET_Customer_ExportUserProfiles
    ' spASPNET_UserProfile_GetProductProfileFromKey
    ' spASPNET_Customer_GetActiveCustomerCodes
    ' spASPNET_Hyster_AddProfile
    ' spASPNET_UserProfile_Add
    ' spASPNET_Hyster_UpdateProfile
    ' spASPNET_UserProfile_Update
    ' spASPNET_UserProfile_UpdateProductProfile
    ' spASPNET_AddEmailToQueue
    ' spASPNET_UserProfile_PromoteUserToSuperUser

    Const PER_CUSTOMER_CONFIGURATION_NONE As Integer = 0
    Const CUSTOMER_HYSTER As Integer = 77
    Const CUSTOMER_YALE As Integer = 680
    Const CUSTOMER_LOVELLS As Integer = 663
    Const CUSTOMER_JUPITER As Integer = 784
    Const CUSTOMER_WURS As Int32 = 579
    Const CUSTOMER_WUIRE As Int32 = 686

    Const CREDIT_LIMIT_ENFORCE_FALSE As Int32 = 0
    Const CREDIT_LIMIT_ENFORCE_TRUE As Int32 = 1

    Dim gnTimeout As Int32
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call GetSiteFeatures()
            Call PageInit()
        End If
        Call InitPerCustomerFormFields(plSelectedCustomerKey)
        Call SetTitle()
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        gnTimeout = Server.ScriptTimeout
        Server.ScriptTimeout = 3600
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        'Server.ScriptTimeout = gnTimeout
        Server.ScriptTimeout = 90
        
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "User Manager"
    End Sub
    
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent5", oConn)
        
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
            WebMsgBox.Show("Error in GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pbProductOwners = CBool(dr("ProductOwners"))
        If dr("UserPermissions") Then
            pbUserPermissions = True
            btnUserGroups.Visible = True
        Else
            btnUserGroups.Visible = False
        End If
        pbProductCredits = dr("ProductCredits")

    End Sub

    Protected Function IsLovells() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsLovells = IIf(gbSiteTypeDefined, gsSiteType = "lovells", nCustomerKey = CUSTOMER_LOVELLS)
    End Function
    
    Protected Function IsJupiterUser() As Boolean
        If pbLoggedOnAsSystemAdministrator Then
            IsJupiterUser = (plSelectedCustomerKey = CUSTOMER_JUPITER)
        Else
            IsJupiterUser = (Session("CustomerKey") = CUSTOMER_JUPITER)
        End If
    End Function

    Protected Function IsWU(ByVal nCustomerKey As Int32) As Boolean
        Dim arrWU() As Integer = {CUSTOMER_WURS, CUSTOMER_WUIRE}
        'Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWU = IIf(gbSiteTypeDefined, gsSiteType = "wu", Array.IndexOf(arrWU, nCustomerKey) >= 0)
    End Function

    Protected Sub PageInit()
        ' NOTE: assumes Product Owners do not have access to User Manager, ie only user types SuperUser or sa have access
        ' if ThisUser is a SuperUser but is not an internal (Sprint) user
        ' then show normal top panel
        ' else show panel that allows choice of customer

        ' If Session("UserType").ToString.Trim.ToLower = "superuser" AndAlso (Not bIsInternalUser()) Then
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not bIsInternalUser() Then
            pbLoggedOnAsSystemAdministrator = False
            plSelectedCustomerKey = Session("CustomerKey")
            Call InitPerCustomerFormFields(plSelectedCustomerKey)
            Call SetAccessLevelChoice()
            Call ShowCustomerUserPanel()
            Call BindUserProfileGrid("UserName")
        Else
            pbLoggedOnAsSystemAdministrator = True
            Call GetCustomerAccountCodes()
            Call SetAccessLevelChoice()
            Call ShowSystemUserPanel()
        End If
        hidCustomer.Value = False
        pbToggleAllowPickAll = True
        pbToggleMaxGrabAll = True
        txtSearchCriteriaAllCustomers.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
        txtSearchCriteriaCustomer.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSSearchUsers.ClientID + "')")
        txtUserProfileProdSeach.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchProducts.ClientID + "')")
        tbNewUserGroupName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnCreateNewUserGroup.ClientID + "')")
        tbRenameUserGroupNewName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnDoRenameUserGroup.ClientID + "')")
    End Sub
    
    Protected Sub InitPerCustomerFormFields(ByVal lCustomerKey As Long)
        lblDepartment.Text = "Cost Centre"
        'txtDepartment.Visible = True
        'ddlUserGroup.Visible = False
        trTelephone.Visible = True
        trHysterDealershipCode.Visible = False ' uses "Title"
        lblCollectionPoint.Text = "Collection Point"
        If lCustomerKey = CUSTOMER_HYSTER Or lCustomerKey = CUSTOMER_YALE Then
            lblDepartment.Text = "Dealership Name"
            trTelephone.Visible = False
            trHysterDealershipCode.Visible = True ' uses "Title"
            lblCollectionPoint.Text = "Nacco Dept Code"
        End If
        trCourierEmailingOptions01.Visible = True
        trCourierEmailingOptions02.Visible = True
        trCourierEmailingOptions03.Visible = True
        trCourierEmailingOptions04.Visible = True
        trCourierEmailingOptions05.Visible = True
        trUserPublicationOptions01.Visible = False
        trUserPublicationOptions02.Visible = False
        trUserPublicationOptions03.Visible = False
        trUserPublicationOptions04.Visible = False
        trUserPublicationOptions05.Visible = False

        If lCustomerKey = CUSTOMER_LOVELLS Then
            trCourierEmailingOptions01.Visible = False
            trCourierEmailingOptions02.Visible = False
            trCourierEmailingOptions03.Visible = False
            trCourierEmailingOptions04.Visible = False
            trCourierEmailingOptions05.Visible = False
            trUserPublicationOptions01.Visible = True
            trUserPublicationOptions02.Visible = True
            'trUserPublicationOptions03.Visible = True
            trUserPublicationOptions04.Visible = True
            'trUserPublicationOptions05.Visible = True
        End If
    End Sub
    
    Protected Function bIsInternalUser() As Boolean
        bIsInternalUser = False
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
            If Not IsDBNull(oDataReader("Customer")) Then
                bIsInternalUser = Not CBool(oDataReader("Customer"))
            Else
                bIsInternalUser = True
            End If
            oDataReader.Close()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        pbIsInternalUser = bIsInternalUser
    End Function
    
    Protected Sub btn_GoBackToProfile_click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        Call ShowAddEditUserPanel()
    End Sub
    
    Protected Sub btn_ReturnToProfile_Click(ByVal s As Object, ByVal e As EventArgs)
        Call ShowAddEditUserPanel()
    End Sub
    
    Protected Sub btn_SaveUserProfileChanges_click(ByVal s As Object, ByVal e As EventArgs)
        If IsValid Then
            Call SaveUserProfileChanges()
        End If
    End Sub
    
    Protected Sub btn_SearchUsers_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call BindUserProfileGrid("UserName")
    End Sub
    
    Protected Sub SearchAllUsers()
        pbIsViewingAllUsers = False
        plSelectedCustomerKey = -1
        lblCustomerName.Text = "[search users from all customers]"
        ddlCustomerAccountCodes.SelectedIndex = 0
        Call BindUserProfileGrid("CustomerAccountCode")
    End Sub
    
    Protected Sub btnExportUserDetails_click(ByVal sender As Object, ByVal e As EventArgs)
        Call ExportUserProfileDetails()
    End Sub
    
    Protected Sub btnExportUserProductDetails_click(ByVal sender As Object, ByVal e As EventArgs)
        Call ExportProductProfileDetails()
    End Sub

    Protected Sub btn_URP_ShowFullList_Click(ByVal sender As Object, ByVal e As EventArgs)
        txtUserProfileProdSeach.Text = ""
        lblProductProfileSearchResult.Visible = False
        dgUserProducts.CurrentPageIndex = 0
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub btn_SearchProducts_Click(ByVal sender As Object, ByVal e As EventArgs)
        dgUserProducts.CurrentPageIndex = 0
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub btn_ShowAllProducts_click(ByVal sender As Object, ByVal e As EventArgs)
        txtUserProfileProdSeach.Text = ""
        lblProductProfileSearchResult.Visible = False
        dgUserProducts.CurrentPageIndex = 0
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub btn_UsersProductList_ShowFullList_Click(ByVal sender As Object, ByVal e As EventArgs)
        txtUserProfileProdSeach.Text = ""
        lblProductProfileSearchResult.Visible = False
        dgUserProducts.CurrentPageIndex = 0
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub btn_ReturnToMyPanel_click(ByVal s As Object, ByVal e As EventArgs)
        Call ReturnToMyPanel()
    End Sub
    
    Protected Sub btn_GoBackToSystemUserStart_click(ByVal s As Object, ByVal e As EventArgs)
        Call ShowSystemUserPanel()
    End Sub
    
    Protected Sub btn_AddUser_click(ByVal s As Object, ByVal e As EventArgs)
        Call AddUser()
    End Sub
    
    Protected Sub AddUser()
        lblSystemUserMessage.Text = ""
        lblSuperUserMessage.Text = ""
        If plSelectedCustomerKey = -1 Then
            WebMsgBox.Show("Please select the customer with whom this user should be associated")
        Else
            Call ShowSelectAccessLevel()
        End If
        pnUserGroup = 0

        If pbProductCredits Then
            trProductCreditsStatus1.Visible = True
            trProductCreditsStatus2.Visible = True
            tbProductCreditsStatus.Text = "Product credits are added when you save the user profile. The credits added depend on the User Group to which you assign the user. To view the credits for this user, save the profile then edit it."
        Else
            trProductCreditsStatus1.Visible = False
            trProductCreditsStatus2.Visible = False
        End If

    End Sub
    
    Protected Sub btn_ContinueToAddUser_click(ByVal s As Object, ByVal e As EventArgs)
        pbIsEditingUser = False
        plSelectedUserKey = -1
        If pbLoggedOnAsSystemAdministrator Then
            lblCustomer.Text = lblCustomerName.Text
        Else
            lblCustomer.Text = Session("CustomerName")
        End If
    
        AddDefaultValuesForNewUser()
        ShowAddEditUserPanel()
    End Sub
    
    Protected Sub ddlCustomerAccountCodes_changed(ByVal s As Object, ByVal e As EventArgs)
        pbIsViewingAllUsers = False
        txtSearchCriteriaAllCustomers.Text = ""
        If IsNumeric(ddlCustomerAccountCodes.SelectedItem.Value) Then
    
            plSelectedCustomerKey = CLng(ddlCustomerAccountCodes.SelectedItem.Value)
            Call InitPerCustomerFormFields(plSelectedCustomerKey)
            
            Call SetAccessLevelChoice()
    
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Customer_GetNameFromKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
    
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.NVarChar, 20)
            paramCustomerKey.Value = plSelectedCustomerKey
            oCmd.Parameters.Add(paramCustomerKey)
    
            Dim paramCustomerName As SqlParameter = New SqlParameter("@CustomerName", SqlDbType.NVarChar, 50)
            paramCustomerName.Direction = ParameterDirection.Output
            oCmd.Parameters.Add(paramCustomerName)
    
            Try
                oConn.Open()
                oCmd.ExecuteNonQuery()
                oConn.Close()
    
                If Not IsDBNull(paramCustomerName.Value) Then
                    lblCustomerName.Text = CStr(paramCustomerName.Value)
                Else
                    lblCustomerName.Text = "[SYSTEM]"
                End If
    
            Catch ex As SqlException
                lblError.Text = ""
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            
            If rbSystemUserCustomerName.Checked Then
                rbSystemUserUserID.Checked = True
            End If
            rbSystemUserCustomerName.Visible = False
            
            BindUserProfileGrid("UserName")
        Else
            plSelectedCustomerKey = -1
            lblCustomerName.Text = ""
            rbSystemUserCustomerName.Visible = True
        End If
    End Sub
    
    Protected Sub ShowAllUsers()
        pbIsViewingAllUsers = True
        plSelectedCustomerKey = -1
        txtSearchCriteriaAllCustomers.Text = ""
        lblCustomerName.Text = "[all users from all customers]"
        ddlCustomerAccountCodes.SelectedIndex = 0
        BindUserProfileGrid("CustomerAccountCode")
    End Sub
    
    Protected Sub btnShowAllCustomerProfiles_click(ByVal s As Object, ByVal e As EventArgs)
        txtSearchCriteriaCustomer.Text = ""
        BindUserProfileGrid("UserName")
    End Sub
    
    Protected Sub btn_ReturnToStart(ByVal s As Object, ByVal e As EventArgs)
        ReturnToStart()
    End Sub
    
    Protected Sub btn_ToggleAllowPick_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        If pbToggleAllowPickAll = True Then
            For Each dgi In dgUserProducts.Items
                cb = CType(dgi.Cells(8).Controls(1), CheckBox)
                cb.Checked = True
            Next dgi
            pbToggleAllowPickAll = False
            b.Text = "deselect all"
        ElseIf pbToggleAllowPickAll = False Then
            For Each dgi In dgUserProducts.Items
                cb = CType(dgi.Cells(8).Controls(1), CheckBox)
                cb.Checked = False
            Next dgi
            pbToggleAllowPickAll = True
            b.Text = "select all"
        End If
    End Sub
    
    Protected Sub btnToggleMaxGrabButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = sender
        If pbToggleMaxGrabAll = True Then
            For Each dgi In dgUserProducts.Items
                cb = CType(dgi.Cells(9).Controls(1), CheckBox)
                cb.Checked = True
            Next dgi
            pbToggleMaxGrabAll = False
            b.Text = "deselect all"
        ElseIf pbToggleMaxGrabAll = False Then
            For Each dgi In dgUserProducts.Items
                cb = CType(dgi.Cells(9).Controls(1), CheckBox)
                cb.Checked = False
            Next dgi
            pbToggleMaxGrabAll = True
            b.Text = "select all"
        End If
    End Sub
    
    Protected Sub btn_ApplyMaxGrabQty_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim tb As TextBox
        If IsNumeric(txtDefaultGrabQty.Text) Then
            For Each dgi In dgUserProducts.Items
                tb = CType(dgi.Cells(10).Controls(1), TextBox)
                tb.Text = txtDefaultGrabQty.Text
            Next dgi
        End If
    End Sub
    
    Protected Sub btn_SaveUserProductProfileChanges_click(ByVal s As Object, ByVal e As EventArgs)
        Call SaveProductProfileChanges()
    End Sub
    
    Protected Sub rblSaUser_IndexChanged(ByVal s As Object, ByVal e As EventArgs)
        psNewUserType = rblSaUser.SelectedValue
        If psNewUserType = "SuperUser" Or psNewUserType.ToLower.Contains("owner") Or psNewUserType = "User" Then
            hidCustomer.Value = True
        Else
            hidCustomer.Value = False  ' Account Handler type requested
        End If
    End Sub
    
    Protected Sub rblSuperUser_IndexChanged(ByVal s As Object, ByVal e As EventArgs)
        psNewUserType = rblSuperUser.SelectedValue
        hidCustomer.Value = True
    End Sub
    
    Protected Sub rblInternalSuperUserWithProductOwner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        psNewUserType = rblInternalSuperUserWithProductOwner.SelectedValue
        hidCustomer.Value = True
        If psNewUserType.Contains("Account") Then
            psNewUserType = "SuperUser"
            hidCustomer.Value = False
        End If
    End Sub

    Protected Sub rblInternalSuperUser_IndexChanged(ByVal s As Object, ByVal e As EventArgs)
        psNewUserType = rblInternalSuperUser.SelectedValue
        hidCustomer.Value = True
        If psNewUserType.Contains("Account") Then
            psNewUserType = "SuperUser"
            hidCustomer.Value = False
        End If
    End Sub
    
    Protected Sub rblSuperUserWithProductOwner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        psNewUserType = rblSuperUserWithProductOwner.SelectedValue
        hidCustomer.Value = True
    End Sub

    Protected Sub GetUserPublicationStatus()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetPublicationProfile", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = plSelectedUserKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("ReceiveMyPublicationOrderAlerts")) Then
                cbUserPublicationsOwnBookingAlerts.Checked = oDataReader("ReceiveMyPublicationOrderAlerts")
            End If
            If Not IsDBNull(oDataReader("ReceiveMyPublicationInactivityAlerts")) Then
                cbUserPublicationsOwnInactivityAlerts.Checked = oDataReader("ReceiveMyPublicationInactivityAlerts")
            End If
            If Not IsDBNull(oDataReader("ReceiveAllPublicationOrderAlerts")) Then
                'cbUserPublicationsAllBookingAlerts.Checked = oDataReader("ReceiveAllPublicationOrderAlerts")
            End If
            If Not IsDBNull(oDataReader("ReceiveAllPublicationInactivityAlerts")) Then
                'cbUserPublicationsAllInactivityAlerts.Checked = oDataReader("ReceiveAllPublicationInactivityAlerts")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetUserPublicationStatus: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function SaveUserPublicationStatus() As Integer
        Dim nReceiveOwnPublicationOrderAlerts As Integer = 0
        Dim nReceiveOwnPublicationInactivityAlerts As Integer = 0
        Dim nReceiveAllPublicationOrderAlerts As Integer = 0
        Dim nReceiveAllPublicationInactivityAlerts As Integer = 0
        
        If cbUserPublicationsOwnBookingAlerts.Checked Then
            nReceiveOwnPublicationOrderAlerts = 1
        End If
        If cbUserPublicationsOwnInactivityAlerts.Checked Then
            nReceiveOwnPublicationInactivityAlerts = 1
        End If
        If cbUserPublicationsAllBookingAlerts.Checked Then
            'nReceiveAllPublicationOrderAlerts = 1
        End If
        If cbUserPublicationsAllInactivityAlerts.Checked Then
            ' nReceiveAllPublicationInactivityAlerts = 1
        End If
        If ExecuteQueryToDataTable("SELECT * FROM UserPublicationProfile WHERE UserKey = " & plSelectedUserKey).Rows.Count > 0 Then
            SaveUserPublicationStatus = ExecuteQueryToDataTable("UPDATE UserPublicationProfile SET ReceiveMyPublicationOrderAlerts = " & nReceiveOwnPublicationOrderAlerts & ", ReceiveMyPublicationInactivityAlerts = " & nReceiveOwnPublicationInactivityAlerts & ", ReceiveAllPublicationOrderAlerts = " & nReceiveAllPublicationOrderAlerts & ", ReceiveAllPublicationInactivityAlerts = " & nReceiveAllPublicationInactivityAlerts & " WHERE UserKey = " & plSelectedUserKey & " SELECT @@ROWCOUNT").Rows(0).Item(0)
        Else
            SaveUserPublicationStatus = ExecuteQueryToDataTable("INSERT INTO UserPublicationProfile (UserKey, ReceiveMyPublicationOrderAlerts, ReceiveMyPublicationInactivityAlerts, ReceiveAllPublicationOrderAlerts, ReceiveAllPublicationInactivityAlerts) VALUES (" & plSelectedUserKey & ", " & nReceiveOwnPublicationOrderAlerts & ", " & nReceiveOwnPublicationInactivityAlerts & ", " & nReceiveAllPublicationOrderAlerts & ", " & nReceiveAllPublicationInactivityAlerts & ") SELECT @@ROWCOUNT").Rows(0).Item(0)
        End If
    End Function
    
    Protected Sub Edit_User(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "Properties" Then
            Dim sUserName As String
            Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
            pbIsEditingUser = True
    
            Dim tcUserKey As TableCell = e.Item.Cells(0)
            plSelectedUserKey = CLng(tcUserKey.Text)
            Dim tcCustomerKey As TableCell = e.Item.Cells(1)
            If IsNumeric(tcCustomerKey.Text) Then
                plSelectedCustomerKey = CLng(tcCustomerKey.Text)
            Else
                plSelectedCustomerKey = 0
            End If
    
            Dim tcUserName As TableCell = e.Item.Cells(4)
            sUserName = tcUserName.Text
            lblProdProfileUserName.Text = sUserName
    
            If plSelectedUserKey > 0 Then
                Dim oDataReader As SqlDataReader
                Dim oConn As New SqlConnection(gsConn)
                'Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey4", oConn)
                Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey5", oConn)
                oCmd.CommandType = CommandType.StoredProcedure
                Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
                oParam.Value = plSelectedUserKey
                Try
                    oConn.Open()
                    oDataReader = oCmd.ExecuteReader()
                    oDataReader.Read()
                    
                    If Not IsDBNull(oDataReader("Customer")) Then
                        hidCustomer.Value = CBool(oDataReader("Customer"))
                    Else
                        hidCustomer.Value = False
                    End If
                    
                    Dim nCustomerKey As Integer = 0
                    If Not IsDBNull(oDataReader("CustomerKey")) Then
                        nCustomerKey = oDataReader("CustomerKey")
                    End If
                    Call InitPerCustomerFormFields(nCustomerKey)
                    
                    If Not IsDBNull(oDataReader("UserId")) Then
                        txtUserId.Text = oDataReader("UserId")
                    End If
                    
                    Dim sPassword As String = String.Empty
                    If Not IsDBNull(oDataReader("Password")) Then
                        sPassword = oPassword.Decrypt(oDataReader("Password"))
                    End If

                    If pbIsInternalUser AndAlso CBool(hidCustomer.Value) = False Then
                        lblInternalUser.Visible = True
                    Else
                        lblInternalUser.Visible = False
                    End If
                    
                    '                    If Not pbIsInternalUser AndAlso CBool(hidCustomer.Value) = False Then     ' external user looking at internal user details - hide password
                    If (Not pbIsInternalUser AndAlso CBool(hidCustomer.Value) = False) Or (Not pbIsInternalUser AndAlso Session("CustomerKey") = CUSTOMER_JUPITER) Then     ' external user looking at internal user details - hide passwordtxtPassword.Text = "****"
                        txtPassword.Text = "****"
                        txtPassword.Enabled = False
                        cbForcePasswordChange.Enabled = False
                    Else
                        txtPassword.Text = sPassword
                        txtPassword.Enabled = True
                        cbForcePasswordChange.Enabled = True
                    End If

                    If Not IsDBNull(oDataReader("FirstName")) Then
                        txtFirstName.Text = oDataReader("FirstName")
                    End If
                    If Not IsDBNull(oDataReader("LastName")) Then
                        txtLastName.Text = oDataReader("LastName")
                    End If

                    If plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE Then
                        If Not IsDBNull(oDataReader("Title")) Then
                            txtDealershipCode.Text = oDataReader("Title")
                        End If
                    End If

                    If Not IsDBNull(oDataReader("Department")) Then
                        txtDepartment.Text = oDataReader("Department")
                    End If
                    
                    If pbUserPermissions Then
                        Dim nThisCustomerKey As Integer
                        If pbLoggedOnAsSystemAdministrator Then
                            nThisCustomerKey = plSelectedCustomerKey
                        Else
                            nThisCustomerKey = Session("CustomerKey")
                        End If
                        If bCustomerHasOneOrMoreUserGroups(nThisCustomerKey) Then
                            trUserGroup.Visible = True
                            Call PopulateUserGroupDropdown(nThisCustomerKey)
                            If Not IsDBNull(oDataReader("UserGroup")) Then
                                pnUserGroup = oDataReader("UserGroup")
                                For i As Integer = 0 To ddlUserGroup.Items.Count - 1
                                    If ddlUserGroup.Items(i).Value = pnUserGroup Then
                                        ddlUserGroup.SelectedIndex = i
                                        Exit For
                                    End If
                                Next
                            Else
                                pnUserGroup = 0
                            End If
                        End If
                    End If
                    
                    If Not IsDBNull(oDataReader("EmailAddr")) Then
                        txtEmailAddr.Text = oDataReader("EmailAddr")
                    End If
                    If Not IsDBNull(oDataReader("Telephone")) Then
                        txtTelephone.Text = oDataReader("Telephone")
                    End If
                    If Not IsDBNull(oDataReader("CollectionPoint")) Then
                        txtCollectionPoint.Text = oDataReader("CollectionPoint")
                    End If
                    If Not IsDBNull(oDataReader("Type")) Then
                        txtAccessLevel.Text = oDataReader("Type")
                    End If
                    
                    If txtAccessLevel.Text.ToLower = "user" Or txtAccessLevel.Text.ToLower.Contains("owner") Then
                        lnkbtnPromoteToSuperUser.Visible = True
                        aAccessLevelHelp.Visible = True
                    Else
                        lnkbtnPromoteToSuperUser.Visible = False
                        aAccessLevelHelp.Visible = False
                    End If
                    
                    If Not IsDBNull(oDataReader("Status")) Then
                        Select Case oDataReader("Status")
                            Case "Active"
                                btnlst_UserStatus.SelectedIndex = 0
                            Case "Suspended"
                                btnlst_UserStatus.SelectedIndex = 1
                        End Select
                    End If
                    If Not IsDBNull(oDataReader("RunningHeaderImage")) Then
                        txtRunningHeaderImage.Text = oDataReader("RunningHeaderImage")
                    End If
                    If Not IsDBNull(oDataReader("CustomerName")) Then
                        lblCustomer.Text = oDataReader("CustomerName")
                    Else
                        lblCustomer.Text = "[SYSTEM]"
                    End If
                    If Not IsDBNull(oDataReader("AbleToViewStock")) Then
                        If oDataReader("AbleToViewStock") Then
                            chkAbleToViewStock.Checked = True
                        Else
                            chkAbleToViewStock.Checked = False
                        End If
                    Else
                        chkAbleToViewStock.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("AbleToCreateStockBooking")) Then
                        If oDataReader("AbleToCreateStockBooking") Then
                            chkAbleToCreateStockBookings.Checked = True
                        Else
                            chkAbleToCreateStockBookings.Checked = False
                        End If
                    Else
                        chkAbleToCreateStockBookings.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("AbleToCreateCollectionRequest")) Then
                        If oDataReader("AbleToCreateCollectionRequest") Then
                            chkAbleToCreateCollectionRequest.Checked = True
                        Else
                            chkAbleToCreateCollectionRequest.Checked = False
                        End If
                    Else
                        chkAbleToCreateCollectionRequest.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("AbleToViewGlobalAddressBook")) Then
                        If oDataReader("AbleToViewGlobalAddressBook") Then
                            chkViewGlobalAddressBook.Checked = True
                        Else
                            chkViewGlobalAddressBook.Checked = False
                        End If
                    Else
                        chkViewGlobalAddressBook.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("AbleToEditGlobalAddressBook")) Then
                        If oDataReader("AbleToEditGlobalAddressBook") Then
                            chkEditGlobalAddressBook.Checked = True
                        Else
                            chkEditGlobalAddressBook.Checked = False
                        End If
                    Else
                        chkEditGlobalAddressBook.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("StockBookingAlert")) Then
                        If oDataReader("StockBookingAlert") Then
                            chkStockBookingAlert.Checked = True
                        Else
                            chkStockBookingAlert.Checked = False
                        End If
                    Else
                        chkStockBookingAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("StockBookingAlertAll")) Then
                        If oDataReader("StockBookingAlertAll") Then
                            chkStockBookingAlertAll.Checked = True
                        Else
                            chkStockBookingAlertAll.Checked = False
                        End If
                    Else
                        chkStockBookingAlertAll.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("StockArrivalAlert")) Then
                        If oDataReader("StockArrivalAlert") Then
                            chkStockArrivalAlert.Checked = True
                        Else
                            chkStockArrivalAlert.Checked = False
                        End If
                    Else
                        chkStockArrivalAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("LowStockAlert")) Then
                        If oDataReader("LowStockAlert") Then
                            chkLowStockAlert.Checked = True
                        Else
                            chkLowStockAlert.Checked = False
                        End If
                    Else
                        chkLowStockAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("ConsignmentBookingAlert")) Then
                        If oDataReader("ConsignmentBookingAlert") Then
                            chkCourierBookingAlert.Checked = True
                        Else
                            chkCourierBookingAlert.Checked = False
                        End If
                    Else
                        chkCourierBookingAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("ConsignmentBookingAlertAll")) Then
                        If oDataReader("ConsignmentBookingAlertAll") Then
                            chkCourierBookingAlertAll.Checked = True
                        Else
                            chkCourierBookingAlertAll.Checked = False
                        End If
                    Else
                        chkCourierBookingAlertAll.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("ConsignmentDespatchAlert")) Then
                        If oDataReader("ConsignmentDespatchAlert") Then
                            chkAWBDespatchAlert.Checked = True
                        Else
                            chkAWBDespatchAlert.Checked = False
                        End If
                    Else
                        chkAWBDespatchAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("ConsignmentDeliveryAlert")) Then
                        If oDataReader("ConsignmentDeliveryAlert") Then
                            chkAWBDeliveryAlert.Checked = True
                        Else
                            chkAWBDeliveryAlert.Checked = False
                        End If
                    Else
                        chkAWBDeliveryAlert.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("MustChangePassword")) Then
                        If oDataReader("MustChangePassword") Then
                            cbForcePasswordChange.Checked = True
                        Else
                            cbForcePasswordChange.Checked = False
                        End If
                    Else
                        cbForcePasswordChange.Checked = False
                    End If
                    If Not IsDBNull(oDataReader("CannotChangePassword")) Then
                        If oDataReader("CannotChangePassword") Then
                            cbUserCannotChangePassword.Checked = True
                        Else
                            cbUserCannotChangePassword.Checked = False
                        End If
                    Else
                        cbUserCannotChangePassword.Checked = False
                    End If
                    oDataReader.Close()
                    If nCustomerKey = CUSTOMER_LOVELLS Then
                        Call GetUserPublicationStatus()
                    End If
                    If pbProductCredits Then
                        trProductCreditsStatus1.Visible = True
                        trProductCreditsStatus2.Visible = True
                        tbProductCreditsStatus.Text = GetProductCreditsStatus()
                    Else
                        trProductCreditsStatus1.Visible = False
                        trProductCreditsStatus2.Visible = False
                    End If
                Catch ex As SqlException
                    lblError.Text = ""
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try

                ShowAddEditUserPanel()
            End If
        ElseIf e.CommandSource.CommandName = "ProductProfile" Then
            cbShowAllowToOrder.Checked = True
            cbShowNotAllowToOrder.Checked = True
            Dim sUserName As String
            Dim sUserLevel As String
            Dim tcUserKey As TableCell = e.Item.Cells(0)
            plSelectedUserKey = CLng(tcUserKey.Text)
            Dim tcUserName As TableCell = e.Item.Cells(4)
            sUserName = tcUserName.Text
            Dim tcUserLevel As TableCell = e.Item.Cells(5)
            sUserLevel = tcUserLevel.Text
            lblProdProfileUserName.Text = sUserName
            dgUserProducts.Visible = False
            tblSaveCancelProductProfile.Visible = False
            lblDefaultMaxGrabQty.Visible = False
            txtDefaultGrabQty.Visible = False
            If sUserLevel.ToLower = "user" Or sUserLevel.ToLower.Contains("owner") Then
                ShowProductProfilePanel()
            ElseIf sUserLevel.ToLower = "superuser" Then
                ShowNoProductProfileMessage()
            End If
        ElseIf e.CommandSource.CommandName = "Email" Then
            Dim sUserName As String
            Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
    
            Dim tcUserKey As TableCell = e.Item.Cells(0)
            plSelectedUserKey = CLng(tcUserKey.Text)
            Dim tcCustomerKey As TableCell = e.Item.Cells(1)
            If IsNumeric(tcCustomerKey.Text) Then
                plSelectedCustomerKey = CLng(tcCustomerKey.Text)
            Else
                plSelectedCustomerKey = 0
            End If
    
            Dim tcUserName As TableCell = e.Item.Cells(4)
            sUserName = tcUserName.Text
            lblProdProfileUserName.Text = sUserName

            If plSelectedUserKey > 0 Then
                Dim oDataReader As SqlDataReader
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey", oConn)
                oCmd.CommandType = CommandType.StoredProcedure
                Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
                oParam.Value = plSelectedUserKey
                Try
                    oConn.Open()
                    oDataReader = oCmd.ExecuteReader()
                    oDataReader.Read()
                    
                    If Not IsDBNull(oDataReader("Customer")) Then
                        hidCustomer.Value = CBool(oDataReader("Customer"))
                    Else
                        hidCustomer.Value = False
                    End If
                    
                    Dim sFirstName As String = "", sLastName As String = ""
                    tbAccessDetailsMessageText.Text = ""
                    
                    If Not IsDBNull(oDataReader("FirstName")) Then
                        sFirstName = oDataReader("FirstName")
                    End If
                    If Not IsDBNull(oDataReader("LastName")) Then
                        sLastName = oDataReader("LastName")
                    End If
                    If sFirstName.Length > 0 Or sLastName.Length > 0 Then
                        tbAccessDetailsMessageText.Text = "Dear " & sFirstName & " " & sLastName & vbNewLine & vbNewLine
                    End If
                    tbAccessDetailsMessageText.Text += "Here is your user ID and password to access the Transworld online system." & vbCrLf
                    
                    If Not IsDBNull(oDataReader("UserId")) Then
                        lblAccessDetailsUserID.Text = oDataReader("UserId")
                    End If
                    If Not IsDBNull(oDataReader("Password")) Then
                        lblAccessDetailsPassword.Text = oPassword.Decrypt(oDataReader("Password"))
                    End If
                    If Not IsDBNull(oDataReader("EmailAddr")) Then
                        tbAccessDetailsEmail.Text = oDataReader("EmailAddr")
                    End If
                    oDataReader.Close()
                Catch ex As SqlException
                    lblError.Text = ""
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
                Call HideAllPanels()
                pnlEmail.Visible = True
            End If

        End If
    End Sub
    
    Protected Function GetProductCreditsStatus() As String
        Dim sSQL As String = "SELECT ProductCode, ProductDescription, StartCredit, RemainingCredit, EnforceCreditLimit, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), CreditStartDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreditStartDateTime, 108)),1,5) AS varchar(20)),'1-jan-2000') 'CreditStartDateTime', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), CreditEndDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreditEndDateTime, 108)),1,5) AS varchar(20)),'1-jan-2000') 'CreditEndDateTime' FROM ProductCredits pcc INNER JOIN LogisticProduct lp ON pcc.LogisticProductKey = lp.LogisticProductKey WHERE pcc.UserKey = " & plSelectedUserKey & " ORDER BY ProductCode"
        Dim sbStatus As New StringBuilder
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        For Each dr As DataRow In dt.Rows
            sbStatus.Append(dr("ProductCode"))
            sbStatus.Append(" (")
            sbStatus.Append(dr("ProductDescription"))
            sbStatus.Append(") ")
            sbStatus.Append(dr("RemainingCredit"))
            sbStatus.Append("/")
            sbStatus.Append(dr("StartCredit"))
            If CBool(dr("EnforceCreditLimit")) Then
                sbStatus.Append(" ENF ")
            Else
                sbStatus.Append(" ODR ")
            End If
            sbStatus.Append(dr("CreditStartDateTime"))
            sbStatus.Append(" - ")
            sbStatus.Append(dr("CreditEndDateTime"))
            sbStatus.Append(Environment.NewLine)
        Next
        GetProductCreditsStatus = sbStatus.ToString
    End Function
    
    Protected Function GetUserGroupForUser(ByVal nUserKey As Integer) As Integer
        Dim o As DataTable = ExecuteQueryToDataTable("SELECT ISNULL(UserGroup,0) FROM UserProfile WHERE UserKey =").Rows(0).Item(0)
        GetUserGroupForUser = 0
    End Function
    
    Protected Sub PopulateUserGroupDropdown(ByVal nCustomerKey As Integer)
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & nCustomerKey & " ORDER BY GroupName", "GroupName", "id")
        ddlUserGroup.Items.Clear()
        ddlUserGroup.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlUserGroup.Items.Add(li)
        Next
    End Sub
    
    Protected Sub BindUserProfileGrid(ByVal SortField As String)
        dgSuperUser.CurrentPageIndex = 0
        dgSystemAdministrator.CurrentPageIndex = 0
        Call PopulateUserProfileGrid(SortField)
    End Sub
    
    Protected Sub RebindUserProfileGrid(ByVal SortField As String)
        Call PopulateUserProfileGrid(SortField)
    End Sub
    
    Protected Function bSetProductProfileButtonVisibility(ByVal DataItem As Object) As String
        bSetProductProfileButtonVisibility = False
        If DataBinder.Eval(DataItem, "Type").ToString.Trim.ToLower = "user" Then
            bSetProductProfileButtonVisibility = True
        End If
    End Function

    Protected Sub PopulateUserProfileGrid(ByVal SortField As String)
        Dim sSearchCriteria As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim sSproc As String
        
        divSystemUserControls.Visible = True
        divSuperUserControls.Visible = True
        
        If cbIncludeSuspendedUsers.Checked Or cbIncludeSuspendedUsers2.Checked Then
            sSproc = "spASPNET_Customer_GetUserProfiles"
        Else
            sSproc = "spASPNET_Customer_GetActiveUserProfiles"
        End If
        Dim oAdapter As New SqlDataAdapter(sSproc, oConn)
        If pbLoggedOnAsSystemAdministrator Then
            sSearchCriteria = txtSearchCriteriaAllCustomers.Text
        Else
            sSearchCriteria = txtSearchCriteriaCustomer.Text
        End If
        
        If sSearchCriteria <> psLastUserSearchString Then
            dgSystemAdministrator.CurrentPageIndex = 0
            dgSuperUser.CurrentPageIndex = 0
            psLastUserSearchString = sSearchCriteria
        End If

        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = plSelectedCustomerKey
            oAdapter.Fill(oDataSet, "Users")
            Dim Source As DataView = oDataSet.Tables("Users").DefaultView
            'Source.Sort = SortField
            
            If Source.Count > 0 Then
                If pbLoggedOnAsSystemAdministrator Then
                    If rbSystemUserUserID.Checked Then
                        Source.Sort = "UserID"
                    ElseIf rbSystemUserUserName.Checked Then
                        Source.Sort = "UserName"
                    Else
                        Source.Sort = "CustomerAccountCode, UserID"
                    End If
                    lblSystemUserMessage.Text = ""
                    dgSystemAdministrator.DataSource = Source
                    dgSystemAdministrator.DataBind()
                    dgSystemAdministrator.Visible = True
                    dgSystemAdministrator.PageSize = ddlUsersPerSystemUserPage.SelectedValue
                    dgSystemAdministrator.CurrentPageIndex = 0
                    'If Source.Count > 10 Then
                    If Source.Count > CInt(ddlUsersPerSystemUserPage.SelectedValue) Then
                        dgSystemAdministrator.PagerStyle.Visible = True
                    Else
                        dgSystemAdministrator.PagerStyle.Visible = False
                    End If
                Else
                    If rbSuperUserUserID.Checked Then
                        Source.Sort = "UserID"
                    Else
                        Source.Sort = "UserName"
                    End If
                    lblSuperUserMessage.Text = ""
                    dgSuperUser.DataSource = Source
                    dgSuperUser.DataBind()
                    dgSuperUser.Visible = True
                    dgSuperUser.PageSize = ddlUserPerSuperUserPage.SelectedValue
                    dgSuperUser.CurrentPageIndex = 0
                    'If Source.Count > 10 Then
                    If Source.Count > CInt(ddlUserPerSuperUserPage.SelectedValue) Then
                        dgSuperUser.PagerStyle.Visible = True
                    Else
                        dgSuperUser.PagerStyle.Visible = False
                    End If
                End If
            Else
                If pbLoggedOnAsSystemAdministrator Then
                    dgSystemAdministrator.Visible = False
                    lblSystemUserMessage.Text = "No matching records"
                Else
                    dgSuperUser.Visible = False
                    lblSuperUserMessage.Text = "No matching records"
                End If
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
            Call ShowDatabaseError()
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub BindUserProductProfileGrid(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetProductProfileFromKey_ByAbleToPick2", oConn)
        
        
        Dim sSearchCriteria As String = txtUserProfileProdSeach.Text

        If sSearchCriteria <> psLastProductSearchString Then
            dgUserProducts.CurrentPageIndex = 0
            psLastProductSearchString = sSearchCriteria
        End If

        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        Else
            cbShowAllowToOrder.Checked = True
            cbShowNotAllowToOrder.Checked = True
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@UserKey").Value = plSelectedUserKey
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@AbleToPick", SqlDbType.Bit))
            If cbShowAllowToOrder.Checked = False And cbShowNotAllowToOrder.Checked = False Then
                cbShowAllowToOrder.Checked = True
                cbShowNotAllowToOrder.Checked = True
            End If
            If cbShowAllowToOrder.Checked = True And cbShowNotAllowToOrder.Checked = True Then
                oAdapter.SelectCommand.Parameters("@AbleToPick").Value = DBNull.Value
            ElseIf cbShowAllowToOrder.Checked Then
                oAdapter.SelectCommand.Parameters("@AbleToPick").Value = 1
            Else
                oAdapter.SelectCommand.Parameters("@AbleToPick").Value = 0
            End If
            oAdapter.Fill(oDataSet, "Products")
            Dim Source As DataView = oDataSet.Tables("Products").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                lblProductProfileMessage.Text = ""
                dgUserProducts.DataSource = Source
                dgUserProducts.DataBind()
                dgUserProducts.Visible = True
                tblSaveCancelProductProfile.Visible = True
                lblDefaultMaxGrabQty.Visible = True
                txtDefaultGrabQty.Visible = True
                lblProductProfileSearchResult.Visible = False
            Else
                dgUserProducts.Visible = False
                tblSaveCancelProductProfile.Visible = False
                lblDefaultMaxGrabQty.Visible = False
                txtDefaultGrabQty.Visible = False
                lblProductProfileSearchResult.Visible = True
            End If
        Catch ex As SqlException
            lblProductProfileMessage.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ExportUserProfileDetails()
        Dim sSearchCriteria As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_ExportUserProfiles", oConn)
        If pbLoggedOnAsSystemAdministrator Then
            sSearchCriteria = txtSearchCriteriaAllCustomers.Text
        Else
            sSearchCriteria = txtSearchCriteriaCustomer.Text
        End If
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = plSelectedCustomerKey
            oAdapter.Fill(oDataSet, "Users")
            Dim Source As DataView = oDataSet.Tables("Users").DefaultView
            Source.Sort = "UserName"
            If Source.Count > 0 Then
    
                Response.Clear()
                'Response.ContentType = "Application/x-msexcel"
                Response.ContentType = "text/csv"
                Response.AddHeader("Content-Disposition", "attachment; filename=user_details.csv")
    
                Dim r As DataRowView
                Dim c As DataColumn
                Dim sItem As String
    
                Dim IgnoredItems As New ArrayList
    
                IgnoredItems.Add("UserKey")
                IgnoredItems.Add("CurrentEncryptedPassword")
    
                For Each c In Source.Table.Columns
                    If Not IgnoredItems.Contains(c.ColumnName) Then
                        If c.ColumnName = "Title" And (plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE) Then
                            Response.Write("NACCO Location Code")
                        ElseIf c.ColumnName = "CollectionPoint" And (plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE) Then
                            Response.Write("NACCO Department Code")
                        Else : Response.Write(c.ColumnName)
                        End If
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
    
                For Each r In Source
                    For Each c In Source.Table.Columns
                        If Not IgnoredItems.Contains(c.ColumnName) Then
                            sItem = (r(c.ColumnName).ToString)
                            sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                            sItem = ControlChars.Quote & sItem & ControlChars.Quote
                            Response.Write(sItem)
                            Response.Write(",")
                        End If
                    Next
                    Response.Write(vbCrLf)
                Next
                Response.End()
            Else
                ' NO MATCHING RECORDS
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ExportProductProfileDetails()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetProductProfileFromKey", oConn)
        Dim sSearchCriteria As String = txtUserProfileProdSeach.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@UserKey").Value = plSelectedUserKey
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.Fill(oDataSet, "Products")
            Dim Source As DataView = oDataSet.Tables("Products").DefaultView
            Source.Sort = "ProductCode"
            If Source.Count > 0 Then
    
                Response.Clear()
                Response.ContentType = "text/csv"
                'Response.ContentType = "Application/x-msexcel"
                Response.AddHeader("Content-Disposition", "attachment; filename=user_product_details.csv")
    
                Dim r As DataRowView
                Dim c As DataColumn
                Dim sItem As String
    
                Dim IgnoredItems As New ArrayList
    
                IgnoredItems.Add("Key")
                IgnoredItems.Add("LogisticProductKey")
    
                For Each c In Source.Table.Columns
                    If Not IgnoredItems.Contains(c.ColumnName) Then
                        Response.Write(c.ColumnName)
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
    
                For Each r In Source
                    For Each c In Source.Table.Columns
                        If Not IgnoredItems.Contains(c.ColumnName) Then
                            sItem = (r(c.ColumnName).ToString)
                            sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                            sItem = ControlChars.Quote & sItem & ControlChars.Quote
                            Response.Write(sItem)
                            Response.Write(",")
                        End If
                    Next
                    Response.Write(vbCrLf)
                Next
    
                Response.End()
    
            Else
                ' NO MATCHING RECORDS
            End If
        Catch ex As SqlException
            lblProductProfileMessage.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub dgSystemAdministrator_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dgSystemAdministrator.CurrentPageIndex = e.NewPageIndex
        Call RebindUserProfileGrid("UserName")
    End Sub
    
    Protected Sub dgSuperUser_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dgSuperUser.CurrentPageIndex = e.NewPageIndex
        Call RebindUserProfileGrid("UserName")
    End Sub

    Protected Sub SortUsersProductsGrid(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        dgUserProducts.CurrentPageIndex = 0
        psSortValue = E.SortExpression
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub SetAccessLevelChoice()
        pnlsaUser.Visible = False
        pnlInternalSuperUser.Visible = False
        pnlInternalSuperUserEnhanced.Visible = False
        pnlSuperUser.Visible = False
        pnlSuperUserWithProductOwner.Visible = False

        If plSelectedCustomerKey = 0 Then
            Select Case Session("UserType")
                Case "sa"
                    pnlsaUser.Visible = True
                Case Else
                    If pbIsInternalUser Then
                        pnlInternalSuperUser.Visible = True
                    Else
                        If pbProductOwners Then
                            pnlSuperUserWithProductOwner.Visible = True
                        Else
                            pnlSuperUser.Visible = True
                        End If
                    End If
            End Select
        Else                       ' note that sa does not offer Product Owner as an option for L&G
            If pbIsInternalUser Then
                If pbProductOwners Then
                    pnlInternalSuperUserEnhanced.Visible = True
                Else
                    pnlInternalSuperUser.Visible = True
                End If
            Else
                If pbProductOwners Then
                    pnlSuperUserWithProductOwner.Visible = True
                Else
                    pnlSuperUser.Visible = True
                End If
            End If
        End If
    End Sub
    
    Protected Sub GetCustomerAccountCodes()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSprocName As String = ""
        If cbOnlyListAccountsWithProducts.Checked Then
            sSprocName = "spASPNET_Customer_GetActiveWithProductsCustomerCodes"
        Else
            sSprocName = "spASPNET_Customer_GetActiveCustomerCodes"
        End If
        ddlCustomerAccountCodes.Items.Clear()
        Dim oCmd As New SqlCommand(sSprocName, oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            ddlCustomerAccountCodes.DataSource = oCmd.ExecuteReader()
            ddlCustomerAccountCodes.DataTextField = "CustomerAccountCode"
            ddlCustomerAccountCodes.DataValueField = "CustomerKey"
            ddlCustomerAccountCodes.DataBind()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub AddDefaultValuesForNewUser()
        txtAccessLevel.Text = psNewUserType
        If txtAccessLevel.Text = "SuperUser" Then
            chkAbleToCreateStockBookings.Checked = True
            chkViewGlobalAddressBook.Checked = True
            chkEditGlobalAddressBook.Checked = True
            chkStockBookingAlert.Checked = True
            chkStockBookingAlertAll.Checked = True
            chkStockArrivalAlert.Checked = True
            chkLowStockAlert.Checked = True
            cbUserPublicationsOwnBookingAlerts.Checked = True
            'cbUserPublicationsAllBookingAlerts.Checked = True
            cbUserPublicationsOwnInactivityAlerts.Checked = True
            'cbUserPublicationsAllInactivityAlerts.Checked = True
        ElseIf txtAccessLevel.Text.ToLower.Contains("owner") Then
            chkAbleToCreateStockBookings.Checked = True
            chkViewGlobalAddressBook.Checked = True
            chkEditGlobalAddressBook.Checked = True
            chkStockBookingAlert.Checked = True
            chkStockBookingAlertAll.Checked = True
            chkStockArrivalAlert.Checked = True
            chkLowStockAlert.Checked = True
            cbUserPublicationsOwnBookingAlerts.Checked = True
            'cbUserPublicationsAllBookingAlerts.Checked = True
            cbUserPublicationsOwnInactivityAlerts.Checked = True
            ' cbUserPublicationsAllInactivityAlerts.Checked = True
        ElseIf txtAccessLevel.Text = "User" Then
            chkAbleToCreateStockBookings.Checked = True
            chkStockBookingAlert.Checked = True
            cbUserPublicationsOwnBookingAlerts.Checked = True
            cbUserPublicationsAllBookingAlerts.Checked = False
            cbUserPublicationsOwnInactivityAlerts.Checked = True
            cbUserPublicationsAllInactivityAlerts.Checked = False
        End If
        btnlst_UserStatus.SelectedIndex = 0
        lnkbtnPromoteToSuperUser.Visible = False
        aAccessLevelHelp.Visible = False
    End Sub
    
    Protected Sub AddNewUser()
        If IsValid Then
            Dim bError As Boolean
            lblError.Text = ""
            lblDBError.Text = String.Empty
            txtUserId.Text = txtUserId.Text.Trim
            If txtUserId.Text.ToLower = "sa" Then
                lblError.Text = "SA is a reserved User ID, please reselect"
                Exit Sub
            End If
            Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
            Dim oConn As New SqlConnection(gsConn)
            
            Dim sStoredProcedure As String
            If plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE Then
                sStoredProcedure = "spASPNET_Hyster_AddProfile4"
            Else
                sStoredProcedure = "spASPNET_UserProfile_Add5"
            End If
            
            Dim oCmd As SqlCommand = New SqlCommand(sStoredProcedure, oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = Session("UserKey") 'User making the add
            oCmd.Parameters.Add(paramUserKey)
            Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 100)
            paramUserId.Value = txtUserId.Text
            oCmd.Parameters.Add(paramUserId)
            Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
            paramPassword.Value = oPassword.Encrypt(txtPassword.Text)
            oCmd.Parameters.Add(paramPassword)
            Dim paramFirstName As SqlParameter = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
            paramFirstName.Value = txtFirstName.Text
            oCmd.Parameters.Add(paramFirstName)
            Dim paramLastName As SqlParameter = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
            paramLastName.Value = txtLastName.Text
            oCmd.Parameters.Add(paramLastName)
            Dim paramTitle As SqlParameter = New SqlParameter("@Title", SqlDbType.NVarChar, 50)
            If plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE Then
                paramTitle.Value = txtDealershipCode.Text
            Else
                paramTitle.Value = Nothing
            End If
            oCmd.Parameters.Add(paramTitle)
            Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
            paramDepartment.Value = txtDepartment.Text
            oCmd.Parameters.Add(paramDepartment)
            
            Dim paramUserGroup As SqlParameter = New SqlParameter("@UserGroup", SqlDbType.Int)
            If pbUserPermissions AndAlso ddlUserGroup.Visible Then
                paramUserGroup.Value = ddlUserGroup.SelectedValue
            Else
                paramUserGroup.Value = 0
            End If
            oCmd.Parameters.Add(paramUserGroup)

            Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
            paramStatus.Value = btnlst_UserStatus.SelectedItem.Text
            oCmd.Parameters.Add(paramStatus)
            Dim paramType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
            paramType.Value = txtAccessLevel.Text
            oCmd.Parameters.Add(paramType)
            Dim paramCustomer As SqlParameter = New SqlParameter("@Customer", SqlDbType.Bit)
            
            ' !!! hidCustomer is empty on second add
            
            If txtAccessLevel.Text = "User" Or txtAccessLevel.Text.ToLower.Contains("owner") _
              OrElse (txtAccessLevel.Text = "SuperUser" And Not hidCustomer.Value = False) _
            Then
                paramCustomer.Value = 1
            Else
                paramCustomer.Value = 0
            End If
            oCmd.Parameters.Add(paramCustomer)
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            If pbLoggedOnAsSystemAdministrator Then
                paramCustomerKey.Value = plSelectedCustomerKey
            Else
                paramCustomerKey.Value = Session("CustomerKey")
            End If
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
            paramEmailAddr.Value = txtEmailAddr.Text
            oCmd.Parameters.Add(paramEmailAddr)
            
            If Not (plSelectedCustomerKey = CUSTOMER_HYSTER Or plSelectedCustomerKey = CUSTOMER_YALE) Then
                Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
                paramTelephone.Value = txtTelephone.Text
                oCmd.Parameters.Add(paramTelephone)
            End If
            
            Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
            paramCollectionPoint.Value = txtCollectionPoint.Text
            oCmd.Parameters.Add(paramCollectionPoint)
            Dim paramURL As SqlParameter = New SqlParameter("@URL", SqlDbType.NVarChar, 100)
            paramURL.Value = "default"
            oCmd.Parameters.Add(paramURL)

            Dim paramAbleToViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit)
            If chkAbleToViewStock.Checked Then
                paramAbleToViewStock.Value = 1
            Else
                paramAbleToViewStock.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToViewStock)

            Dim paramAbleToCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit)
            If chkAbleToCreateStockBookings.Checked Then
                paramAbleToCreateStockBooking.Value = 1
            Else
                paramAbleToCreateStockBooking.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToCreateStockBooking)

            Dim paramAbleToCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit)
            If chkAbleToCreateCollectionRequest.Checked Then
                paramAbleToCreateCollectionRequest.Value = 1
            Else
                paramAbleToCreateCollectionRequest.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToCreateCollectionRequest)
            Dim paramAbleToViewGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit)
            If chkViewGlobalAddressBook.Checked Then
                paramAbleToViewGlobalAddressBook.Value = 1
            Else
                paramAbleToViewGlobalAddressBook.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToViewGlobalAddressBook)
            Dim paramAbleToEditGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit)
            If chkEditGlobalAddressBook.Checked Then
                paramAbleToEditGlobalAddressBook.Value = 1
            Else
                paramAbleToEditGlobalAddressBook.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToEditGlobalAddressBook)
            Dim paramRunningHeader As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
            If txtRunningHeaderImage.Text <> "" Then
                paramRunningHeader.Value = txtRunningHeaderImage.Text
            Else
                paramRunningHeader.Value = "default"
            End If
            oCmd.Parameters.Add(paramRunningHeader)
            Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit)
            If chkStockBookingAlert.Checked Then
                paramStockBookingAlert.Value = 1
            Else
                paramStockBookingAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramStockBookingAlert)
            Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit)
            If chkStockBookingAlertAll.Checked Then
                paramStockBookingAlertAll.Value = 1
            Else
                paramStockBookingAlertAll.Value = 0
            End If
            oCmd.Parameters.Add(paramStockBookingAlertAll)
            Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit)
            If chkStockArrivalAlert.Checked Then
                paramStockArrivalAlert.Value = 1
            Else
                paramStockArrivalAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramStockArrivalAlert)
            Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit)
            If chkLowStockAlert.Checked Then
                paramLowStockAlert.Value = 1
            Else
                paramLowStockAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramLowStockAlert)
            Dim paramCourierBookingAlert As SqlParameter = New SqlParameter("@ConsignmentBookingAlert", SqlDbType.Bit)
            If chkCourierBookingAlert.Checked Then
                paramCourierBookingAlert.Value = 1
            Else
                paramCourierBookingAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierBookingAlert)
            Dim paramCourierBookingAlertAll As SqlParameter = New SqlParameter("@ConsignmentBookingAlertAll", SqlDbType.Bit)
            If chkCourierBookingAlertAll.Checked Then
                paramCourierBookingAlertAll.Value = 1
            Else
                paramCourierBookingAlertAll.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierBookingAlertAll)
            Dim paramCourierDespatchAlert As SqlParameter = New SqlParameter("@ConsignmentDespatchAlert", SqlDbType.Bit)
            If chkAWBDespatchAlert.Checked Then
                paramCourierDespatchAlert.Value = 1
            Else
                paramCourierDespatchAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierDespatchAlert)
            Dim paramCourierDeliveryAlert As SqlParameter = New SqlParameter("@ConsignmentDeliveryAlert", SqlDbType.Bit)
            If chkAWBDeliveryAlert.Checked Then
                paramCourierDeliveryAlert.Value = 1
            Else
                paramCourierDeliveryAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierDeliveryAlert)
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
                plSelectedUserKey = CLng(oCmd.Parameters("@UserProfileKey").Value)
                Dim nForcePasswordChange As Int32
                If cbForcePasswordChange.Checked Then
                    nForcePasswordChange = 1
                End If
                Call ExecuteQueryToDataTable("UPDATE UserProfile SET MustChangePassword = " & nForcePasswordChange.ToString & " WHERE [key] = " & plSelectedUserKey.ToString)
            Catch ex As SqlException
                bError = True
                oTrans.Rollback("AddRecord")
                If ex.Number = 2627 Then
                    WebMsgBox.Show("'" & txtUserId.Text & "' - this User ID is already taken. User IDs must be unique. Please choose another ID.")
                    lblError.Text = "This User ID is already taken. Please select another User ID."
                    Exit Sub
                Else
                    lblDBError.Text = ex.ToString
                End If
            Finally
                oConn.Close()
            End Try
            If bError Then
                ShowDatabaseError()
            Else
                If plSelectedCustomerKey = CUSTOMER_LOVELLS Then
                    Call SaveUserPublicationStatus()
                End If

                WebMsgBox.Show("User details added")
            End If
            If pbUserPermissions AndAlso ddlUserGroup.Visible AndAlso txtAccessLevel.Text.ToLower = "user" Then
                If ddlUserGroup.SelectedValue <> pnUserGroup Then
                    Call SetPermissionsForUser(plSelectedUserKey, ddlUserGroup.SelectedValue)
                End If
            End If
        End If
    End Sub
    
    Protected Sub SetPermissionsForUser(ByVal nDestinationUserKey As Long, ByVal nUserGroup As Integer)
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT TOP 1 [key] FROM UserProfile WHERE [key] <> " & nDestinationUserKey & " AND Type = 'User' AND UserGroup  = " & nUserGroup)
        If oDataTable.Rows.Count = 1 Then
            Dim nSourceUserKey As Integer = oDataTable.Rows(0).Item(0)
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProductProfileCloneFromUser", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramSourceUserKey As SqlParameter = New SqlParameter("@SourceUserKey", SqlDbType.Int)
            paramSourceUserKey.Value = nSourceUserKey
            oCmd.Parameters.Add(paramSourceUserKey)
            Dim paramDestinationUserKey As SqlParameter = New SqlParameter("@DestinationUserKey", SqlDbType.Int)
            paramDestinationUserKey.Value = nDestinationUserKey
            oCmd.Parameters.Add(paramDestinationUserKey)
            Try
                oConn.Open()
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                WebMsgBox.Show("Error in SetPermissionsForUser: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        Else
            WebMsgBox.Show("No existing user found from which to copy permissions! This user still has default permissioning.")
        End If
    End Sub
    
    Protected Sub ReturnToMyPanel()
        lblError.Text = ""
        pnlAddEditUser.EnableViewState = False
        plSelectedUserKey = -1
        If pbLoggedOnAsSystemAdministrator Then
            ShowSystemUserPanel()
            BindUserProfileGrid("CustomerAccountCode")
        Else
            ShowCustomerUserPanel()
            BindUserProfileGrid("UserName")
        End If
    End Sub
    
    Protected Sub ReturnToStart()
        If pbLoggedOnAsSystemAdministrator Then
            ShowSystemUserPanel()
        Else
            ShowCustomerUserPanel()
        End If
    End Sub
    
    Protected Function nGetCustomerKeyForUser(ByVal nUserKey As Integer) As Integer
        nGetCustomerKeyForUser = 0
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
            If Not IsDBNull(oDataReader("CustomerKey")) Then
                nGetCustomerKeyForUser = CInt(oDataReader("CustomerKey"))
            End If
            oDataReader.Close()
        Catch ex As SqlException
            WebMsgBox.Show("Error in nGetCustomerKeyForUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub RemoveProductCredits()
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE UserKey = " & plSelectedUserKey)
    End Sub
    
    Protected Sub AddProductCredits()
        Call SetTemplateChangeFlag()
        Exit Sub
        
        Dim nCurrentUserGroup As Int32 = ExecuteQueryToDataTable("SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & plSelectedUserKey).Rows(0).Item(0)
        Dim dtProductCreditControl As DataTable = ExecuteQueryToDataTable("SELECT LogisticProductKey, Credit, NextRefreshDateTime, EnforceCreditLimit, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits FROM ProductCreditControl WHERE UserOrUserGroup = " & nCurrentUserGroup)
        For Each drProductCreditControlEntry In dtProductCreditControl.Rows
            Dim nLogisticProductKey As Int32 = drProductCreditControlEntry("LogisticProductKey")
            Dim nCredit As Int32 = drProductCreditControlEntry("Credit")
            Dim nEnforceCreditLimit As Int32 = drProductCreditControlEntry("EnforceCreditLimit")
            'Dim dateNextRefreshDateTime As DateTime = drProductCreditControlEntry("NextRefreshDateTime")
            Dim sRefreshDaysOrMonths As String = drProductCreditControlEntry("RefreshDaysOrMonths")
            Dim nRefreshInterval As Int32 = drProductCreditControlEntry("RefreshInterval")
            Dim sCarryOverCredit As String = drProductCreditControlEntry("CarryOverCredit")
            Dim nMaxCredits As Int32 = drProductCreditControlEntry("MaxCredits")

            Dim sEndDateExpression As String
            If sRefreshDaysOrMonths.ToLower = "d" Then
                sEndDateExpression = "DATEADD(DAY, " & nRefreshInterval & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "')"
            Else
                sEndDateExpression = "DATEADD(MONTH, " & nRefreshInterval & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "')"
            End If

            'Call ExecuteQueryToDataTable("INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, CreditStartDateTime, CreditEndDateTime) VALUES (" & nLogisticProductKey & ", " & plSelectedUserKey & ", " & nCreditAmount & ", " & nCreditAmount & ", GETDATE(), '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "')")
            Call ExecuteQueryToDataTable("INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) VALUES (" & nLogisticProductKey & ", " & plSelectedUserKey & ", " & nCredit & ", " & nCredit & ", " & nEnforceCreditLimit & ", GETDATE(), " & sEndDateExpression & ")")
        Next
    End Sub
    
    Protected Sub SetTemplateChangeFlag()
        WriteRegistry(Registry.LocalMachine, "SOFTWARE\CourierSoftware\ProductCredits", "TemplateChange", "true", RegistryValueKind.String)
    End Sub
    
    Protected Sub WriteRegistry(ByVal rkRegistryKey As RegistryKey, ByVal sSubKey As String, ByVal sKey As String, ByVal sValue As String, ByVal rvkType As RegistryValueKind)
        Try
            rkRegistryKey.OpenSubKey(name:=sSubKey, writable:=True).SetValue(name:=sKey, value:=sValue, valueKind:=rvkType)
        Catch e As Exception
            WebMsgBox.Show("Error writing registry: " & e.Message)
            'Globals.log.debug("Failed to write to registry key '" + tree.Name + "\" + subKey + "\" + subKey + "'. " + e.Message)
        End Try
    End Sub

    Protected Sub PermissionJupiterPODProducts()
        Dim sSQL As String = "UPDATE UserProductProfile SET AbleToView = 1, AbleToPick = 1, AbleToEdit = 0, AbleToArchive = 0, AbleToDelete = 0, ApplyMaxGrab = 1, MaxGrabQty = 1000 WHERE UserKey = " & plSelectedUserKey & " AND ProductKey IN (SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_JUPITER & " AND ISNULL(Misc2, 0) > 0)"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub SaveUserProfileChanges()
        Dim nCustomerKey As Integer
        lblDBError.Text = String.Empty
        lblError.Text = ""
        If plSelectedUserKey = -1 Then
            Call AddNewUser()
            nCustomerKey = nGetCustomerKeyForUser(plSelectedUserKey)
            If IsWU(nCustomerKey) Then     ' this is a kludge that must be sorted out if Product Credits are used for non WU users; necessary because ProdCreds are linked to the URL, not the customer
                Call AddProductCredits()
            End If
            If IsJupiterUser() Then
                Call PermissionJupiterPODProducts()
            End If
            pbIsEditingUser = False
        Else
            Dim nCurrentUserGroup As Int32 = 0
            If pbProductCredits Then
                nCurrentUserGroup = ExecuteQueryToDataTable("SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & plSelectedUserKey).Rows(0).Item(0)
            End If
            nCustomerKey = nGetCustomerKeyForUser(plSelectedUserKey)
            Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
            Dim oConn As New SqlConnection(gsConn)
            Dim oTrans As SqlTransaction

            Dim sStoredProcedure As String
            If nCustomerKey = CUSTOMER_HYSTER Or nCustomerKey = CUSTOMER_YALE Then
                sStoredProcedure = "spASPNET_Hyster_UpdateProfile5"
            Else
                sStoredProcedure = "spASPNET_UserProfile_Update5"
            End If
            
            Dim oCmd As SqlCommand = New SqlCommand(sStoredProcedure, oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = CLng(Session("UserKey"))
            oCmd.Parameters.Add(paramUserKey)
            Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
            paramUserProfileKey.Value = plSelectedUserKey
            oCmd.Parameters.Add(paramUserProfileKey)
            Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.NVarChar, 100)
            paramUserId.Value = txtUserId.Text
            oCmd.Parameters.Add(paramUserId)
            Dim paramPassword As SqlParameter = New SqlParameter("@Password", SqlDbType.NVarChar, 24)
            paramPassword.Value = oPassword.Encrypt(txtPassword.Text)
            oCmd.Parameters.Add(paramPassword)
            Dim paramFirstName As SqlParameter = New SqlParameter("@FirstName", SqlDbType.NVarChar, 50)
            paramFirstName.Value = txtFirstName.Text
            oCmd.Parameters.Add(paramFirstName)
            Dim paramLastName As SqlParameter = New SqlParameter("@LastName", SqlDbType.NVarChar, 50)
            paramLastName.Value = txtLastName.Text
            oCmd.Parameters.Add(paramLastName)
            Dim paramTitle As SqlParameter = New SqlParameter("@Title", SqlDbType.NVarChar, 50)

            If nCustomerKey = CUSTOMER_HYSTER Or nCustomerKey = CUSTOMER_YALE Then
                paramTitle.Value = txtDealershipCode.Text
            Else
                paramTitle.Value = Nothing
            End If
            oCmd.Parameters.Add(paramTitle)

            Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
            paramDepartment.Value = String.Empty
            'If ddlUserGroup.SelectedIndex > 0 Then
            '    paramDepartment.Value = ddlUserGroup.SelectedItem.Text
            'End If
            paramDepartment.Value = txtDepartment.Text.Trim
            oCmd.Parameters.Add(paramDepartment)

            Dim paramUserGroup As SqlParameter = New SqlParameter("@UserGroup", SqlDbType.Int)
            If pbUserPermissions AndAlso ddlUserGroup.Visible Then
                paramUserGroup.Value = ddlUserGroup.SelectedValue
            Else
                paramUserGroup.Value = 0
            End If
            oCmd.Parameters.Add(paramUserGroup)

            Dim paramType As SqlParameter = New SqlParameter("@Type", SqlDbType.NVarChar, 20)
            paramType.Value = txtAccessLevel.Text
            oCmd.Parameters.Add(paramType)
            Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.NVarChar, 20)
            paramStatus.Value = btnlst_UserStatus.SelectedItem.Text
            oCmd.Parameters.Add(paramStatus)
            Dim paramCustomer As SqlParameter = New SqlParameter("@Customer", SqlDbType.Bit)
            If txtAccessLevel.Text = "User" OrElse txtAccessLevel.Text.ToLower.Contains("owner") _
              OrElse (txtAccessLevel.Text = "SuperUser" And Not hidCustomer.Value = False) _
            Then
                paramCustomer.Value = 1
            Else
                paramCustomer.Value = 0
            End If
            oCmd.Parameters.Add(paramCustomer)
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            If pbLoggedOnAsSystemAdministrator Then
                paramCustomerKey.Value = plSelectedCustomerKey
            Else
                paramCustomerKey.Value = Session("CustomerKey")
            End If
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
            paramEmailAddr.Value = txtEmailAddr.Text
            oCmd.Parameters.Add(paramEmailAddr)

            If Not (nCustomerKey = CUSTOMER_HYSTER Or nCustomerKey = CUSTOMER_YALE) Then
                Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
                paramTelephone.Value = txtTelephone.Text
                oCmd.Parameters.Add(paramTelephone)
            End If

            Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
            paramCollectionPoint.Value = txtCollectionPoint.Text
            oCmd.Parameters.Add(paramCollectionPoint)
            Dim paramURL As SqlParameter = New SqlParameter("@URL", SqlDbType.NVarChar, 100)
            paramURL.Value = "Default"
            oCmd.Parameters.Add(paramURL)
            
            Dim paramAbleToViewStock As SqlParameter = New SqlParameter("@AbleToViewStock", SqlDbType.Bit, 1)
            If chkAbleToViewStock.Checked Then
                paramAbleToViewStock.Value = 1
            Else
                paramAbleToViewStock.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToViewStock)
            
            Dim paramAbleToCreateStockBooking As SqlParameter = New SqlParameter("@AbleToCreateStockBooking", SqlDbType.Bit, 1)
            If chkAbleToCreateStockBookings.Checked Then
                paramAbleToCreateStockBooking.Value = 1
            Else
                paramAbleToCreateStockBooking.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToCreateStockBooking)
            Dim paramAbleToCreateCollectionRequest As SqlParameter = New SqlParameter("@AbleToCreateCollectionRequest", SqlDbType.Bit, 1)
            If chkAbleToCreateCollectionRequest.Checked Then
                paramAbleToCreateCollectionRequest.Value = 1
            Else
                paramAbleToCreateCollectionRequest.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToCreateCollectionRequest)
            Dim paramAbleToViewGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToViewGlobalAddressBook", SqlDbType.Bit, 1)
            If chkViewGlobalAddressBook.Checked Then
                paramAbleToViewGlobalAddressBook.Value = 1
            Else
                paramAbleToViewGlobalAddressBook.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToViewGlobalAddressBook)
            Dim paramAbleToEditGlobalAddressBook As SqlParameter = New SqlParameter("@AbleToEditGlobalAddressBook", SqlDbType.Bit, 1)
            If chkEditGlobalAddressBook.Checked Then
                paramAbleToEditGlobalAddressBook.Value = 1
            Else
                paramAbleToEditGlobalAddressBook.Value = 0
            End If
            oCmd.Parameters.Add(paramAbleToEditGlobalAddressBook)
            Dim paramRunningHeaderImage As SqlParameter = New SqlParameter("@RunningHeaderImage", SqlDbType.NVarChar, 100)
            If txtRunningHeaderImage.Text <> "" Then
                paramRunningHeaderImage.Value = txtRunningHeaderImage.Text
            Else
                paramRunningHeaderImage.Value = "default"
            End If
            oCmd.Parameters.Add(paramRunningHeaderImage)
            Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit, 1)
            If chkStockBookingAlert.Checked Then
                paramStockBookingAlert.Value = 1
            Else
                paramStockBookingAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramStockBookingAlert)
            Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit, 1)
            If chkStockBookingAlertAll.Checked Then
                paramStockBookingAlertAll.Value = 1
            Else
                paramStockBookingAlertAll.Value = 0
            End If
            oCmd.Parameters.Add(paramStockBookingAlertAll)
            Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit, 1)
            If chkStockArrivalAlert.Checked Then
                paramStockArrivalAlert.Value = 1
            Else
                paramStockArrivalAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramStockArrivalAlert)
            Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit, 1)
            If chkLowStockAlert.Checked Then
                paramLowStockAlert.Value = 1
            Else
                paramLowStockAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramLowStockAlert)
            Dim paramCourierBookingAlert As SqlParameter = New SqlParameter("@ConsignmentBookingAlert", SqlDbType.Bit, 1)
            If chkCourierBookingAlert.Checked Then
                paramCourierBookingAlert.Value = 1
            Else
                paramCourierBookingAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierBookingAlert)
            Dim paramCourierBookingAlertAll As SqlParameter = New SqlParameter("@ConsignmentBookingAlertAll", SqlDbType.Bit, 1)
            If chkCourierBookingAlertAll.Checked Then
                paramCourierBookingAlertAll.Value = 1
            Else
                paramCourierBookingAlertAll.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierBookingAlertAll)
            Dim paramCourierDespatchAlert As SqlParameter = New SqlParameter("@ConsignmentDespatchAlert", SqlDbType.Bit, 1)
            If chkAWBDespatchAlert.Checked Then
                paramCourierDespatchAlert.Value = 1
            Else
                paramCourierDespatchAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierDespatchAlert)
            Dim paramCourierDeliveryAlert As SqlParameter = New SqlParameter("@ConsignmentDeliveryAlert", SqlDbType.Bit, 1)
            If chkAWBDeliveryAlert.Checked Then
                paramCourierDeliveryAlert.Value = 1
            Else
                paramCourierDeliveryAlert.Value = 0
            End If
            oCmd.Parameters.Add(paramCourierDeliveryAlert)
            Dim paramMustChangePassword As SqlParameter = New SqlParameter("@MustChangePassword", SqlDbType.Bit, 1)
            If cbForcePasswordChange.Checked Then
                paramMustChangePassword.Value = 1
            Else
                paramMustChangePassword.Value = 0
            End If
            oCmd.Parameters.Add(paramMustChangePassword)
            Try
                oConn.Open()
                oCmd.Connection = oConn
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                If ex.Number = 2627 Then
                    WebMsgBox.Show("'" & txtUserId.Text & "' - this User ID is already taken. User IDs must be unique. Please choose another ID.")
                    lblError.Text = "This User ID is already taken. User IDs must be unique. Please choose another ID."
                    Exit Sub
                Else
                    lblDBError.Text = ex.ToString
                    Call ShowDatabaseError()
                End If
            Finally
                oConn.Close()
            End Try
            If lblError.Text = String.Empty And lblDBError.Text = String.Empty Then
                If nCustomerKey = CUSTOMER_LOVELLS Then
                    Call SaveUserPublicationStatus()
                End If
            End If
            If pbUserPermissions AndAlso ddlUserGroup.Visible AndAlso txtAccessLevel.Text.ToLower = "user" Then
                If pnUserGroup <> ddlUserGroup.SelectedValue Then
                    Call SetPermissionsForUser(plSelectedUserKey, ddlUserGroup.SelectedValue)
                End If
            End If
            If pbProductCredits Then
                'If nCurrentUserGroup > 0 Then
                If nCurrentUserGroup <> ddlUserGroup.SelectedValue Then
                    Call RemoveProductCredits()
                    Call AddProductCredits()
                End If
                'End If
            End If
            WebMsgBox.Show("User profile updated")
        End If
        Call ReturnToMyPanel()
    End Sub
    
    Protected Sub SetDefaultUserPermissions()
        Dim nDefaultProductGroup As Integer = GetDefaultProductGroup()
        If nDefaultProductGroup > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "INSERT INTO UP_UserPermissions (UserKey, GroupOrProductKey, LastModifiedDateTime, LastUpdateBy) VALUES ("
            sSQL += plSelectedUserKey & ", " & nDefaultProductGroup & ", GETDATE(), " & Session("UserKey") & ")"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show("Error in SetDefaultUserPermissions: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub ClearDefaultUserPermissions()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "DELETE FROM UP_UserPermissions WHERE UserKey = " & plSelectedUserKey
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in ClearDefaultUserPermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ApplyMaxGrabs()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserPermissions_ApplyMaxGrabsForUser", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = plSelectedUserKey
        oCmd.Parameters.Add(paramUserKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in ApplyMaxGrabs: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ApplyUserPermissions()
        Call RefreshPermissions(plSelectedUserKey, 0, 0)
    End Sub

    Protected Sub RefreshPermissions(ByVal sUserKey As String, ByVal nCustomerKey As Integer, ByVal nProductGroupKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserPermissions_Apply2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = nCustomerKey
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = sUserKey
        oCmd.Parameters.Add(paramUserKey)
        Dim paramProductGroupKey As SqlParameter = New SqlParameter("@ProductGroupKey", SqlDbType.Int)
        paramProductGroupKey.Value = nProductGroupKey
        oCmd.Parameters.Add(paramProductGroupKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in RefreshPermissions: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    'Protected Function GetCurrentUserPermissionsGroup() As String
    '    GetCurrentUserPermissionsGroup = String.Empty
    '    Dim oDataReader As SqlDataReader = Nothing
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim sSQL As String = "SELECT ProductGroup FROM UP_UserPermissions up INNER JOIN UP_ProductPermissionGroups ppg ON up.GroupOrProductKey = ppg.[id] WHERE GroupOrProductKey >= 1000000 AND UserKey = " & plSelectedUserKey
    '    Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
    '    Try
    '        oConn.Open()
    '        oDataReader = oCmd.ExecuteReader()
    '        If oDataReader.HasRows Then
    '            oDataReader.Read()
    '            GetCurrentUserPermissionsGroup = oDataReader(0)
    '        End If
    '    Catch ex As Exception
    '        WebMsgBox.Show("Error in GetCurrentUserPermissionsGroup: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Function
    
    Protected Function GetDefaultProductGroup() As Integer
        GetDefaultProductGroup = 0
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim nCustomerKey As Integer = nGetCustomerKeyForUser(plSelectedUserKey)
        Dim sSQL As String = "SELECT id FROM UP_ProductPermissionGroups WHERE DefaultUserGroup = " & ddlUserGroup.SelectedItem.Value & " AND CustomerKey = " & nCustomerKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetDefaultProductGroup = CInt(oDataReader(0))
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetDefaultProductGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SaveProductProfileChanges()
        Const CELL_PRODUCT_PROFILE_KEY As Integer = 0
        Const CELL_ABLE_TO_PICK As Integer = 8
        Const CELL_APPLY_MAX_GRAB As Integer = 9
        Const CELL_MAX_GRAB_QTY As Integer = 10
        
        Dim lProductProfiles As Long
        Dim dgi As DataGridItem
        Dim tcProductProfileKey As TableCell
        Dim lProductProfileKey As Long
        Dim cbAbleToPick As CheckBox
        Dim cbApplyMaxGrab As CheckBox
        Dim tbMaxGrabQty As TextBox
        Dim oConn As New SqlConnection(gsConn)
        Dim nAValidProductProfileKey As Integer
    
        For Each dgi In dgUserProducts.Items
            lProductProfiles = lProductProfiles + 1
            tcProductProfileKey = dgi.Cells(CELL_PRODUCT_PROFILE_KEY)
            lProductProfileKey = CLng(tcProductProfileKey.Text)
            cbAbleToPick = CType(dgi.Cells(CELL_ABLE_TO_PICK).Controls(1), CheckBox)
            cbApplyMaxGrab = CType(dgi.Cells(CELL_APPLY_MAX_GRAB).Controls(1), CheckBox)
            tbMaxGrabQty = CType(dgi.Cells(CELL_MAX_GRAB_QTY).Controls(1), TextBox)
    
            If lProductProfileKey > 0 Then
                nAValidProductProfileKey = lProductProfileKey
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_UpdateProductProfile", oConn)
                oCmd.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@ProductProfileKey", SqlDbType.Int, 4)
                param1.Value = lProductProfileKey
                oCmd.Parameters.Add(param1)
                Dim param2 As SqlParameter = New SqlParameter("@AbleToPick", SqlDbType.Bit)
                If cbAbleToPick.Checked Then
                    param2.Value = 1
                Else
                    param2.Value = 0
                End If
                oCmd.Parameters.Add(param2)
                Dim param3 As SqlParameter = New SqlParameter("@ApplyMaxGrab", SqlDbType.Bit)
                If cbApplyMaxGrab.Checked Then
                    param3.Value = 1
                Else
                    param3.Value = 0
                End If
                oCmd.Parameters.Add(param3)
                Dim param4 As SqlParameter = New SqlParameter("@MaxGrabQty", SqlDbType.Int, 4)
                If IsNumeric(tbMaxGrabQty.Text) Then
                    param4.Value = CLng(tbMaxGrabQty.Text)
                Else
                    param4.Value = 0
                End If
                oCmd.Parameters.Add(param4)
                Try
                    oConn.Open()
                    oCmd.Connection = oConn
                    oCmd.ExecuteNonQuery()
                    lProductProfileKey = 0
                Catch ex As SqlException
                    lblProductProfileMessage.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
            End If
        Next dgi
        Dim sConfirmationMessage As String
        sConfirmationMessage = ""
        If lProductProfiles = 1 Then
            sConfirmationMessage = lProductProfiles.ToString & " product profile has been updated."
        Else
            sConfirmationMessage = lProductProfiles.ToString & " product profiles have been updated."
        End If
        If pbUserPermissions Then
            Dim nUserChangedCount As Integer = CloneUserProfiles(nAValidProductProfileKey)
            If nUserChangedCount > 0 Then
                sConfirmationMessage += " " & nUserChangedCount & " users in the same User Group have been updated."
            End If
        End If
        WebMsgBox.Show(sConfirmationMessage)
    End Sub

    Protected Function CloneUserProfiles(ByVal nProductProfileKey As Long) As Integer
        CloneUserProfiles = 0
        Dim nUserKey As Long = ExecuteQueryToDataTable("SELECT UserKey FROM UserProductProfile WHERE [key] = " & nProductProfileKey).Rows(0).Item(0)
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProductProfileCloneByUserGroup", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SourceUserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SourceUserKey").Value = nUserKey
        oAdapter.SelectCommand.CommandTimeout = 600
        Try
            oAdapter.Fill(oDataTable)
            CloneUserProfiles = oDataTable.Rows(0).Item(0)
        Catch ex As Exception
            WebMsgBox.Show("Error in CloneUserProfiles: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub HideAllPanels()
        pnlAddEditUser.Visible = False
        pnlChooseUserType.Visible = False
        pnlCustomerUser.Visible = False
        pnlDatabaseError.Visible = False
        pnlProductProfile.Visible = False
        pnlShowNoProductProfileMessage.Visible = False
        pnlSystemUser.Visible = False
        pnlEmail.Visible = False
        pnlConfirmPermissionsChange.Visible = False
        pnlUserGroups.Visible = False
        pnlNewUserGroup.Visible = False
        pnlRenameUserGroup.Visible = False
        pnlAdjustProductCredits.Visible = False
    End Sub
    
    Protected Sub ShowDatabaseError()
        Call HideAllPanels()
        pnlDatabaseError.Visible = True
    End Sub
    
    Protected Sub ShowAddEditUserPanel()
        Call HideAllPanels()
        lblInternalUser.Visible = False
        pnlAddEditUser.Visible = True
        If pbUserPermissions = True Then
            If Not pbIsEditingUser Then
                Dim nCustomerKey As Integer
                If pbLoggedOnAsSystemAdministrator Then
                    nCustomerKey = plSelectedCustomerKey
                Else
                    nCustomerKey = Session("CustomerKey")
                End If
                If bCustomerHasOneOrMoreUserGroups(nCustomerKey) Then
                    trUserGroup.Visible = True
                    Call PopulateUserGroupDropdown(nCustomerKey)
                End If
            End If
        End If
    End Sub
    
    Protected Function bCustomerHasOneOrMoreUserGroups(ByVal nCustomerKey As Integer) As Boolean
        bCustomerHasOneOrMoreUserGroups = False
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT TOP 1 GroupName FROM UP_UserPermissionGroups WHERE CustomerKey = " & nCustomerKey)
        If oDataTable.Rows.Count > 0 Then
            bCustomerHasOneOrMoreUserGroups = True
        End If
    End Function
    
    Protected Sub ShowProductProfilePanel()
        Call HideAllPanels()
        If pbLoggedOnAsSystemAdministrator Then
        Else
        End If
        pnlProductProfile.Visible = True
    End Sub
    
    Protected Sub ShowSystemUserPanel()
        Call HideAllPanels()
        pnlSystemUser.Visible = True
        ddlCustomerAccountCodes.Focus()
    End Sub
    
    Protected Sub ShowCustomerUserPanel()
        Call HideAllPanels()
        pnlCustomerUser.Visible = True
    End Sub
    
    Protected Sub ShowSelectAccessLevel()
        Call HideAllPanels()
        pnlChooseUserType.Visible = True
        Call InitSelectAccessLevel()
    End Sub
    
    Protected Sub InitSelectAccessLevel()
        For Each li As ListItem In rblSaUser.Items
            If li.Text.ToLower = ("user") Then
                li.Selected = True
                Exit For
            End If
        Next
        For Each li As ListItem In rblInternalSuperUserWithProductOwner.Items
            If li.Text.ToLower = ("user") Then
                li.Selected = True
                Exit For
            End If
        Next
        For Each li As ListItem In rblInternalSuperUser.Items
            If li.Text.ToLower = ("user") Then
                li.Selected = True
                Exit For
            End If
        Next
        For Each li As ListItem In rblSuperUserWithProductOwner.Items
            If li.Text.ToLower = ("user") Then
                li.Selected = True
                Exit For
            End If
        Next
        For Each li As ListItem In rblSuperUser.Items
            If li.Text.ToLower = ("user") Then
                li.Selected = True
                Exit For
            End If
        Next
        hidCustomer.Value = False
    End Sub
    
    Protected Sub ShowNoProductProfileMessage()
        Call HideAllPanels()
        pnlShowNoProductProfileMessage.Visible = True
    End Sub

    Protected Sub ShowConfirmPermissionsChangePanel()
        Call HideAllPanels()
        pnlConfirmPermissionsChange.Visible = True
    End Sub
    
    Protected Sub btnShowAllUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllUsers()
    End Sub

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SearchAllUsers()
    End Sub

    Protected Sub btnAddNewUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddUser()
    End Sub
    
    Protected Sub cbIncludeSuspendedUsers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindUserProfileGrid("CustomerAccountCode")
    End Sub
    
    Protected Sub btnSendEmail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSubject As String = "Your account details to access the Transworld online system"
        Dim sBody As New StringBuilder
        sBody.Append(tbAccessDetailsMessageText.Text)
        sBody.Append(vbNewLine)
        sBody.Append("User ID:  " & lblAccessDetailsUserID.Text)
        sBody.Append(vbNewLine)
        sBody.Append("Password: " & lblAccessDetailsPassword.Text)
        sBody.Append(vbNewLine)
        sBody.Append(vbNewLine)
        sBody.Append("Please do not reply to this message as this address is not monitored. In case of difficulty please contact your Account Handler.")
        sBody.Replace(vbNewLine, "<br />" & vbNewLine)

        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_AddEmailToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageTypeId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageTypeId").Value = "WEB_USERID_REQUEST"

            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")

            oCmd.Parameters.Add(New SqlParameter("@Recipient", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Recipient").Value = tbAccessDetailsEmail.Text

            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 50))
            oCmd.Parameters("@Subject").Value = sSubject

            oCmd.Parameters.Add(New SqlParameter("@Body", SqlDbType.NVarChar, 3880))
            oCmd.Parameters("@Body").Value = sBody.ToString

            oCmd.Parameters.Add(New SqlParameter("@EmailMessageQueueKey", SqlDbType.Int, 4))
            oCmd.Parameters("@EmailMessageQueueKey").Direction = ParameterDirection.Output

            oConn.Open()
            oCmd.ExecuteNonQuery()

        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call ReturnToMyPanel()
    End Sub

    Protected Sub lnkbtnPromoteToSuperUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PromoteUserToSuperUser()
    End Sub
    
    Protected Function bUserIsActiveProductOwner(ByVal nUserProfileKey As Integer) As Boolean
        bUserIsActiveProductOwner = False
        If ExecuteQueryToDataTable("SELECT * FROM ProductGroup WHERE ProductOwner1 = " & nUserProfileKey & " OR ProductOwner2 = " & nUserProfileKey).Rows.Count > 0 Then
            bUserIsActiveProductOwner = True
        End If
    End Function
    
    Protected Sub PromoteUserToSuperUser()
        If ExecuteQueryToDataTable("SELECT * FROM ProductGroup WHERE ProductOwner1 = " & plSelectedUserKey & " OR ProductOwner2 = " & plSelectedUserKey).Rows.Count > 0 Then
            WebMsgBox.Show("Cannot promote while user is Product Owner (or Deputy Product Owner) for one or more product groups.")
            Exit Sub
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim sStoredProcedure As String
        sStoredProcedure = "spASPNET_UserProfile_PromoteUserToSuperUser2"
        Dim oCmd As SqlCommand = New SqlCommand(sStoredProcedure, oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey")) 'user making the update
        oCmd.Parameters.Add(paramUserKey)
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = plSelectedUserKey
        oCmd.Parameters.Add(paramUserProfileKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            txtAccessLevel.Text = "SuperUser"
        Catch ex As SqlException
            lblDBError.Text = ex.ToString
            Call ShowDatabaseError()
        Finally
            oConn.Close()
        End Try
        WebMsgBox.Show("This user is now a Super User")
    End Sub
    
    Protected Sub dgSuperUser_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dgi As DataGridItem = e.Item
        Dim b As Button
        Dim s As String
        If dgi.ItemType = ListItemType.Item Or dgi.ItemType = ListItemType.AlternatingItem Then
            s = dgi.Cells(5).Text
            b = dgi.Cells(9).FindControl("btnSuperUserProductProfile")
            If s = "SuperUser" Then
                b.Visible = False
            End If
        End If
    End Sub
    
    Protected Sub chkAbleToViewStock_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            chkAbleToCreateStockBookings.Checked = False
        End If
    End Sub

    Protected Sub chkAbleToCreateStockBookings_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            chkAbleToViewStock.Checked = False
        End If
    End Sub
    
    Protected Sub cbOnlyListAccountsWithProducts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetCustomerAccountCodes()
    End Sub
    
    Protected Sub btnContinueToChangePermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearDefaultUserPermissions()
        Call SetDefaultUserPermissions()
        Call ApplyUserPermissions()
        Call ApplyMaxGrabs()
        Call ReturnToMyPanel()
    End Sub

    Protected Sub InitDefinedUserGroupsListbox()
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT * FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        lbDefinedUserGroups.Items.Clear()
        lbDefinedUserGroups.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            lbDefinedUserGroups.Items.Add(li)
        Next
    End Sub
   
    Protected Sub btnNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlNewUserGroup.Visible = True
        tbNewUserGroupName.Text = String.Empty
        tbNewUserGroupName.Focus()
        SetEnableRenameAndRemoveUserGroup(False)
    End Sub
   
    Protected Sub btnRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlRenameUserGroup.Visible = True
        tbRenameUserGroupNewName.Text = lbDefinedUserGroups.SelectedItem.Text
        tbRenameUserGroupNewName.Focus()
    End Sub
   
    Protected Sub btnRemoveUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SetEnableRenameAndRemoveUserGroup(False)
        Call RemoveUserGroup()
        Call InitDefinedUserGroupsListbox()
    End Sub
    
    Protected Sub RemoveUserGroup()
        If CInt(ExecuteQueryToDataTable("SELECT COUNT (*) FROM UserProfile WHERE UserGroup = " & lbDefinedUserGroups.SelectedValue).Rows(0)(0)) > 0 Then
            WebMsgBox.Show("Cannot remove. This user group is still referenced by one or more users.")
        Else
            If Not ExecuteNonQuery("DELETE FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND [id] = " & lbDefinedUserGroups.SelectedValue) Then
                WebMsgBox.Show("Error attempting to remove user group.")
            End If
        End If
    End Sub

    Protected Sub btnCreateNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CreateNewUserGroup()
    End Sub

    Protected Sub CreateNewUserGroup()
        tbNewUserGroupName.Text = tbNewUserGroupName.Text.Trim
        If tbNewUserGroupName.Text = String.Empty Then
            WebMsgBox.Show("Please enter a name.")
        Else
            If UserGroupExists(tbNewUserGroupName.Text) Then
                WebMsgBox.Show("The name '" & tbNewUserGroupName.Text & "' is already in use! Please choose another name.")
                tbNewUserGroupName.Text = String.Empty
                tbNewUserGroupName.Focus()
            Else
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand
                Dim sSQL As String = "INSERT INTO UP_UserPermissionGroups (GroupName, CustomerKey, LastModifiedDateTime, LastUpdateBy) VALUES ('"
                sSQL += tbNewUserGroupName.Text.Replace("'", "''") & "', " & Session("CustomerKey") & ", GETDATE(), " & Session("UserKey") & ")"
                Try
                    oConn.Open()
                    oCmd = New SqlCommand(sSQL, oConn)
                    oCmd.ExecuteNonQuery()
                    WebMsgBox.Show("User group '" & tbNewUserGroupName.Text & "' successfully created.")
                    Call HideAllPanels()
                    Call InitDefinedUserGroupsListbox()
                    pnlUserGroups.Visible = True
                Catch ex As Exception
                    WebMsgBox.Show("Error in CreateNewUserGroup: " & ex.Message)
                Finally
                    oConn.Close()
                End Try
            End If
        End If
    End Sub
   
    Protected Function UserGroupExists(ByVal sUserGroupName As String) As Boolean
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT * FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " AND GroupName = '" & sUserGroupName.Replace("'", "''") & "'", "GroupName", "GroupName")
        If oListItemCollection.Count = 0 Then
            UserGroupExists = False
        Else
            UserGroupExists = True
        End If
    End Function
    
    Protected Sub btnCancelNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub

    Protected Sub btnDoRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RenameUserGroup()
        Call InitDefinedUserGroupsListbox()
        lbDefinedUserGroups.SelectedIndex = -1
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub
   
    Protected Sub RenameUserGroup() ' no longer needed once change made to use id of UP_UserPermissionGroups
        tbRenameUserGroupNewName.Text = tbRenameUserGroupNewName.Text.Trim
        If UserGroupExists(tbRenameUserGroupNewName.Text) Then
            WebMsgBox.Show("This name is already in use! Please choose another name.")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand
            Dim sSQL As String = "UPDATE UP_UserPermissionGroups SET "
            sSQL += "GroupName = '" & tbRenameUserGroupNewName.Text.Replace("'", "''") & "', LastModifiedDateTime = GETDATE(), LastUpdateBy = " & Session("UserKey")
            sSQL += " WHERE CustomerKey = " & Session("CustomerKey") & " AND GroupName = '" & lbDefinedUserGroups.SelectedItem.Text & "'"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                WebMsgBox.Show("Error in RenameUserGroup: " & ex.Message)
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
   
    Protected Sub btnCancelRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlUserGroups.Visible = True
    End Sub

    Protected Sub SetEnableRenameAndRemoveUserGroup(ByVal bStatus As Boolean)
        btnRenameUserGroup.Enabled = bStatus
        btnRemoveUserGroup.Enabled = bStatus
    End Sub
   
    Protected Sub lbDefinedUserGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If lbDefinedUserGroups.Items(0).Text.Contains("please select") Then
            lbDefinedUserGroups.Items.RemoveAt(0)
        End If
        If lbDefinedUserGroups.SelectedIndex >= 0 Then
            Call SetEnableRenameAndRemoveUserGroup(True)
        End If
    End Sub
    
    Protected Sub btnUserGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitUserGroupsPanel()
    End Sub
   
    Protected Sub InitUserGroupsPanel()
        Call HideAllPanels()
        pnlUserGroups.Visible = True
        Call InitDefinedUserGroupsListbox()
    End Sub

    Protected Sub btnBackFromUserGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReturnToMyPanel()
    End Sub

    Protected Sub btnBackFromNewUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitUserGroupsPanel()
    End Sub

    Protected Sub btnBackFromRenameUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitUserGroupsPanel()
    End Sub
    
    Protected Sub btnUsersReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Users Report")
        sbText.Append(Bold("USERS REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report is divided into 4 sections. <b>Section 1</b> shows the defined user groups. <b>Section 2</b> lists the users who are not in any user group. <b>Section 3</b> shows, for each defined user group, the users in that group. <b>Section 4</b> lists Super Users.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<i>NOTE: The permissions system applies to standard Users only. Super Users are automatically permissioned for all products. A list of Super Users is shown in section 4.</i>")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT [id], GroupName FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        sbText.Append(Bold("1. Defined user groups (" & oListItemCollection.Count & ") are:"))
        Call NewLine(sbText)
        For Each liGroupName As ListItem In oListItemCollection
            sbText.Append(liGroupName.Text)
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Call NewLine(sbText)
        Dim oListItemCollection5 As ListItemCollection = ExecuteQueryToListItemCollection("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'User' AND Status = 'Active' AND CustomerKey = " & Session("CustomerKey") & " AND NOT ISNULL(UserGroup,0) IN (SELECT [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & ") ORDER BY FirstName", "UserKey", "UserName")
        sbText.Append("<hr />")
        sbText.Append(Bold("2. Users not in a defined group (" & oListItemCollection5.Count & ")"))
        Call NewLine(sbText)
        For Each liUserName As ListItem In oListItemCollection5
            sbText.Append(liUserName.Value)
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Dim oListItemCollectionA As ListItemCollection = ExecuteQueryToListItemCollection("SELECT GroupName, [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "id")
        For Each liGroupName As ListItem In oListItemCollectionA
            Dim oListItemCollection2 As ListItemCollection
            If liGroupName.Text.ToLower.Contains("suspend") Then
                oListItemCollection2 = ExecuteQueryToListItemCollection("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND UserGroup = " & liGroupName.Value & " ORDER BY FirstName", "UserKey", "UserName")
            Else
                oListItemCollection2 = ExecuteQueryToListItemCollection("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & " AND Status = 'Active' AND UserGroup = " & liGroupName.Value & " ORDER BY FirstName", "UserKey", "UserName")
            End If
            sbText.Append("<hr />")
            sbText.Append(Bold("3. Users in group " & liGroupName.Text & " (" & oListItemCollection2.Count & ")"))
            Call NewLine(sbText)
            For Each liUserName As ListItem In oListItemCollection2
                sbText.Append(liUserName.Value)
                Call NewLine(sbText)
            Next
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Dim oListItemCollection6 As ListItemCollection = ExecuteQueryToListItemCollection("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'SuperUser' AND Status = 'Active' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName", "UserKey", "UserName")
        sbText.Append(Bold("4. Super Users (" & oListItemCollection6.Count & ")"))
        Call NewLine(sbText)
        For Each liUserName As ListItem In oListItemCollection6
            sbText.Append(liUserName.Value)
            Call NewLine(sbText)
        Next
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "UsersReport")
    End Sub

    Protected Function Bold(ByVal sString As String) As String
        Bold = "<b>" & sString & "</b>"
    End Function
   
    Protected Sub NewLine(ByRef sbText As StringBuilder)
        sbText.Append("<br />" & Environment.NewLine)
    End Sub
   
    Protected Sub AddHTMLPreamble(ByRef sbText As StringBuilder, ByVal sTitle As String)
        sbText.Append("<html><head><title>")
        sbText.Append(sTitle)
        sbText.Append("</title><style>")
        sbText.Append("body { font-family: Verdana; font-size : xx-small }")
        sbText.Append("</style></head><body>")
    End Sub
   
    Protected Sub AddHTMLPostamble(ByRef sbText As StringBuilder)
        sbText.Append("</body></html>")
    End Sub
   
    Private Sub ExportData(ByVal sData As String, ByVal sFilename As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & sFilename & ".htm")
        Response.ContentType = "text/html"
   
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sData)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()
        Response.End()
    End Sub

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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection: " & ex.Message)
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

    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("UM_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_ProductOwners") = Value
        End Set
    End Property
    
    Property pbUserPermissions() As Boolean
        Get
            Dim o As Object = ViewState("UM_UserPermissions")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_UserPermissions") = Value
        End Set
    End Property
    
    Property pbIsViewingAllUsers() As Boolean
        Get
            Dim o As Object = ViewState("UM_IsViewingAllUsers")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_IsViewingAllUsers") = Value
        End Set
    End Property
    
    Property pbIsEditingUser() As Boolean
        Get
            Dim o As Object = ViewState("UM_IsEditingUser")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_IsEditingUser") = Value
        End Set
    End Property
    
    Property pbLoggedOnAsSystemAdministrator() As Boolean
        Get
            Dim o As Object = ViewState("UM_IsInSystemMode")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_IsInSystemMode") = Value
        End Set
    End Property
    
    Property pbIsInternalUser() As Boolean
        Get
            Dim o As Object = ViewState("UM_IsInternalUser")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_IsInternalUser") = Value
        End Set
    End Property
    
    Property plSelectedUserKey() As Long
        Get
            Dim o As Object = ViewState("UM_SelectedUserKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UM_SelectedUserKey") = Value
        End Set
    End Property
    
    Property plSelectedCustomerKey() As Long
        Get
            Dim o As Object = ViewState("UM_SelectedCustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UM_SelectedCustomerKey") = Value
        End Set
    End Property
    
    Property pbToggleAllowPickAll() As Boolean
        Get
            Dim o As Object = ViewState("UM_ToggleAllowPickAll")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_ToggleAllowPickAll") = Value
        End Set
    End Property
    
    Property pbToggleMaxGrabAll() As Boolean
        Get
            Dim o As Object = ViewState("UM_ToggleMaxGrabAll")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_ToggleMaxGrabAll") = Value
        End Set
    End Property
    
    Property psNewUserType() As String
        Get
            Dim o As Object = ViewState("UM_NewUserType")
            If o Is Nothing Then
                Return "User"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("UM_NewUserType") = Value
        End Set
    End Property

    Property plPerCustomerConfiguration() As Long
        Get
            Dim o As Object = ViewState("UM_PerCustomerConfiguration")
            If o Is Nothing Then
                Return PER_CUSTOMER_CONFIGURATION_NONE
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("UM_PerCustomerConfiguration") = Value
        End Set
    End Property
    
    Property pnUserGroup() As Integer
        Get
            Dim o As Object = ViewState("UM_UserGroup")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("UM_UserGroup") = Value
        End Set
    End Property

    Property psSortValue() As String
        Get
            Dim o As Object = ViewState("UM_SortValue")
            If o Is Nothing Then
                Return "ProductCode"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("UM_SortValue") = Value
            If Value.ToLower = "productcode" Then
                lblSortValue.Text = "Product Code"
            ElseIf Value.ToLower = "productdate" Then
                lblSortValue.Text = "Product Date"
            ElseIf Value.ToLower = "productdescription" Then
                lblSortValue.Text = "Product Description"
            ElseIf Value.ToLower = "productcategory" Then
                lblSortValue.Text = "Product Category"
            ElseIf Value.ToLower = "subcategory" Then
                lblSortValue.Text = "Product Sub Category"
            End If
        End Set
    End Property

    Property psLastProductSearchString() As String
        Get
            Dim o As Object = ViewState("UM_LastSearchProductString")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("UM_LastSearchProductString") = Value
        End Set
    End Property

    Property psLastUserSearchString() As String
        Get
            Dim o As Object = ViewState("UM_LastSearchUserString")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("UM_LastSearchUserString") = Value
        End Set
    End Property

    Property pbProductCredits() As Boolean
        Get
            Dim o As Object = ViewState("UM_ProductCredits")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UM_ProductCredits") = Value
        End Set
    End Property

    Protected Sub dgUserProducts_PageIndexChanged(ByVal sender As Object, ByVal e As DataGridPageChangedEventArgs)
        dgUserProducts.CurrentPageIndex = e.NewPageIndex
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub ddlUsersPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        dgUserProducts.PageSize = ddl.SelectedValue
        dgUserProducts.CurrentPageIndex = 0
        Call BindUserProductProfileGrid(psSortValue)
    End Sub
    
    Protected Sub ddlUsersPerSystemUserPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSystemAdministrator.PageSize = ddlUsersPerSystemUserPage.SelectedValue
        dgSystemAdministrator.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")
    End Sub

    Protected Sub ddlUserPerSuperUserPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSuperUser.PageSize = ddlUserPerSuperUserPage.SelectedValue
        dgSuperUser.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")
    End Sub
    
    Protected Sub rbSystemUserUserID_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSystemAdministrator.PageSize = ddlUsersPerSystemUserPage.SelectedValue
        dgSystemAdministrator.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")

    End Sub

    Protected Sub rbSystemUserUserName_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSystemAdministrator.PageSize = ddlUsersPerSystemUserPage.SelectedValue
        dgSystemAdministrator.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")

    End Sub

    Protected Sub rbSystemUserCustomerName_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSystemAdministrator.PageSize = ddlUsersPerSystemUserPage.SelectedValue
        dgSystemAdministrator.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")

    End Sub

    Protected Sub rbSuperUserUserID_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSuperUser.PageSize = ddlUserPerSuperUserPage.SelectedValue
        dgSuperUser.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")
    End Sub

    Protected Sub rbSuperUserUserName_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSuperUser.PageSize = ddlUserPerSuperUserPage.SelectedValue
        dgSuperUser.CurrentPageIndex = 0
        Call PopulateUserProfileGrid("UserName")
    End Sub
    
    Protected Sub lnkbtnAdjustCredits_Click(sender As Object, e As System.EventArgs)
        Call HideAllPanels()
        pnlAdjustProductCredits.Visible = True
        lblAdjustProductCreditsUser.Text = txtUserId.Text
        Call AdjustProductCredits()
    End Sub
    
    Protected Sub AdjustProductCredits()
        Call BindProductCredits()
    End Sub

    Protected Sub BindProductCredits()
        Dim sSQL As String = "SELECT [id], ProductCode + ' ' + ProductDate 'Product', ProductDescription, StartCredit, RemainingCredit, EnforceCreditLimit, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), CreditStartDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreditStartDateTime, 108)),1,5) AS varchar(20)),'(never)') 'CreditStartDateTime', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), CreditEndDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), CreditEndDateTime, 108)),1,5) AS varchar(20)),'(never)') 'CreditEndDateTime' FROM ProductCredits pc INNER JOIN LogisticProduct lp ON pc.LogisticProductKey = lp.LogisticProductKey WHERE UserKey = " & plSelectedUserKey
        Dim dtProductCredits As DataTable = ExecuteQueryToDataTable(sSQL)
        gvAdjustProductCredits.DataSource = dtProductCredits
        gvAdjustProductCredits.DataBind()
    End Sub

    Protected Sub lnkbtnEditProductCredit_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        trNewProductCredit.Visible = False
        trEditProductCredit.Visible = True
        btnEditCreditSave.CommandArgument = lnkbtn.CommandArgument
        Dim sSQL As String = "SELECT * FROM ProductCredits WHERE [id] = " & lnkbtn.CommandArgument
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        lblEditCreditProduct.Text = ExecuteQueryToDataTable("SELECT ProductCode + ' ' + ProductDate 'Product' FROM LogisticProduct WHERE LogisticProductKey = " & dr("LogisticProductKey")).Rows(0).Item(0)
        rntbEditCredit.Text = dr("RemainingCredit")
        If dr("EnforceCreditLimit") = CREDIT_LIMIT_ENFORCE_TRUE Then
            cbEditCreditEnforce.Checked = True
        Else
            cbEditCreditEnforce.Checked = False
        End If
        'cbEditCreditEnforce.Checked = CBool(dr("EnforceCreditLimit"))
        rdtpEditCreditStartDateTime.SelectedDate = dr("CreditStartDateTime")
        rdtpEditCreditEndDateTime.SelectedDate = dr("CreditEndDateTime")
        rntbEditCredit.Focus
    End Sub

    Protected Sub lnkbtnRemoveProductCredit_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE [id] = " & lnkbtn.CommandArgument)
        Call BindProductCredits()
    End Sub

    Protected Sub btnNewProductCredit_Click(sender As Object, e As System.EventArgs)
        Dim btn As Button = sender
        rntbNewCredit.Text = String.Empty
        rdtpNewCreditStartDateTime.SelectedDate = Date.Today
        rdtpNewCreditEndDateTime.SelectedDate = Date.Today
        cbNewCreditEnforce.Checked = True
        
        trEditProductCredit.Visible = False
        Dim sSQL As String = "SELECT DISTINCT SUBSTRING(ProductCode + ' ' + ProductDescription, 1, 40) 'Product', LogisticProductKey FROM LogisticProduct lp INNER JOIN UserProductProfile upp ON lp.LogisticProductKey = upp.ProductKey WHERE ArchiveFlag = 'N' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " AND LogisticProductKey NOT IN (SELECT LogisticProductKey FROM ProductCredits WHERE UserKey = " & plSelectedUserKey & ") AND upp.AbleToPick = 1 ORDER BY Product, LogisticProductKey"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            ddlNewCreditProduct.Items.Add(New ListItem("- please select -", 0))
            For Each dr As DataRow In dt.Rows
                ddlNewCreditProduct.Items.Add(New ListItem(dr("Product"), dr("LogisticProductKey")))
            Next
        
            ddlNewCreditProduct.SelectedIndex = 0
            ddlNewCreditProduct.Focus()
            trNewProductCredit.Visible = True
        Else
            WebMsgBox.Show("All products available to this user already have credits applied.\n\nPlease edit the existing product credits, or delete the credits for the product of interest and then re-apply them.")
        End If
    End Sub
    
    Protected Sub btnEditCreditSave_Click(sender As Object, e As System.EventArgs)
        Dim btn As Button = sender
        Dim sCredit As String = rntbEditCredit.Text.Trim
        If Not IsNumeric(sCredit) Then
            WebMsgBox.Show("Please specify the credit required.")
            Exit Sub
        End If
        If rdtpEditCreditStartDateTime.SelectedDate Is Nothing Then
            WebMsgBox.Show("Please specify the start date required.")
            Exit Sub
        End If
        If rdtpEditCreditEndDateTime.SelectedDate Is Nothing Then
            WebMsgBox.Show("Please specify the start date required.")
            Exit Sub
        End If
        Dim sStartDateTime As String = rdtpEditCreditStartDateTime.SelectedDate
        Dim sEndDateTime As String = rdtpEditCreditEndDateTime.SelectedDate
        Dim nEnforceCreditLimit As Int32
        If cbEditCreditEnforce.Checked Then
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_TRUE
        Else
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_FALSE
        End If

        Dim sSQL As String = "UPDATE ProductCredits SET StartCredit = " & sCredit & ", RemainingCredit = " & sCredit & ", EnforceCreditLimit = " & nEnforceCreditLimit & ", CreditStartDateTime = '" & Date.Parse(rdtpEditCreditStartDateTime.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "', CreditEndDateTime = '" & Date.Parse(rdtpEditCreditEndDateTime.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "' WHERE [id] = " & btn.CommandArgument
        Call ExecuteQueryToDataTable(sSQL)
        Call BindProductCredits()
        trEditProductCredit.Visible = False
    End Sub

    Protected Sub btnEditCreditCancel_Click(sender As Object, e As System.EventArgs)
        trEditProductCredit.Visible = False
    End Sub

    Protected Sub btnNewCreditSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlNewCreditProduct.SelectedIndex <= 0 Then
            WebMsgBox.Show("Please specify the product to which credits should be applied.")
            Exit Sub
        End If
        Dim sCredit As String = rntbNewCredit.Text.Trim
        If Not IsNumeric(sCredit) Then
            WebMsgBox.Show("Please specify the number of credits required.")
            Exit Sub
        End If
        Dim nEnforceCreditLimit As Int32
        If cbNewCreditEnforce.Checked Then
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_TRUE
        Else
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_FALSE
        End If

        If rdtpNewCreditStartDateTime.SelectedDate Is Nothing Then
            WebMsgBox.Show("Please specify the start date required.\n\nNormally you should use the current date and time.")
            Exit Sub
        End If
        If rdtpNewCreditEndDateTime.SelectedDate Is Nothing Then
            WebMsgBox.Show("Please specify the end date required.")
            Exit Sub
        End If
        
        If rdtpNewCreditStartDateTime.SelectedDate >= rdtpNewCreditEndDateTime.SelectedDate Then
            WebMsgBox.Show("Start date must be before end date.")
            Exit Sub
        End If
        'Dim sStartDateTime As String = rdtpNewCreditStartDateTime.SelectedDate
        'Dim sEndDateTime As String = rdtpNewCreditEndDateTime.SelectedDate

        ' need to check start date is before end date (but later)
        If ddlNewCreditProduct.SelectedIndex = 0 Then
            WebMsgBox.Show("Please specify the product to which you want to apply credits.")
            Exit Sub
        End If
        trNewProductCredit.Visible = False
        Dim sSQL As String = "INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) VALUES (" & ddlNewCreditProduct.SelectedValue & ", " & plSelectedUserKey & ", " & rntbNewCredit.Text & ", " & rntbNewCredit.Text & ", " & nEnforceCreditLimit & ", '" & Date.Parse(rdtpNewCreditStartDateTime.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "', '" & Date.Parse(rdtpNewCreditEndDateTime.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "')"
        Call ExecuteQueryToDataTable(sSQL)
        Call BindProductCredits()
    End Sub

    Protected Sub btnNewCreditCancel_Click(sender As Object, e As System.EventArgs)
        trNewProductCredit.Visible = False
    End Sub
    
    Protected Function gvProductCreditType(ByVal DataItem As Object) As String
        If CBool(DataBinder.Eval(DataItem, "EnforceCreditLimit")) Then
            gvProductCreditType = "ENFORCED"
        Else
            gvProductCreditType = "OVERDRAFT"
        End If
    End Function

    Protected Sub btnBackFromAdjustProductCredits_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReturnToMyPanel()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>User Manager</title>
    <style type="text/css">
.RadInput_Default{font:12px "segoe ui",arial,sans-serif}.RadInput{vertical-align:middle;width:160px}.RadInput_Default{font:12px "segoe ui",arial,sans-serif}.RadInput{vertical-align:middle;width:160px}.RadInput_Default{font:12px "segoe ui",arial,sans-serif}.RadInput{vertical-align:middle;width:160px}</style>
</head>
<body>
    <form id="frmUserManager" runat="Server">
    <main:Header ID="ctlHeader" runat="server"></main:Header>
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_usermanager">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlSystemUser" runat="server" Visible="False" Width="100%">
        <table style="width: 100%; font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td align="left" valign="middle" style="white-space: nowrap; height: 27px;">
                    <asp:DropDownList ID="ddlCustomerAccountCodes" runat="server" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlCustomerAccountCodes_changed" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                    <asp:CheckBox ID="cbOnlyListAccountsWithProducts" runat="server" Checked="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="accts with products only" AutoPostBack="True" OnCheckedChanged="cbOnlyListAccountsWithProducts_CheckedChanged" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnShowAllUsers" runat="server" Text="show all users"
                        OnClick="btnShowAllUsers_Click" />
                    &nbsp;&nbsp;&nbsp;<asp:Label ID="Label13" Font-Size="XX-Small" Font-Names="Verdana"
                        runat="server" Text="search" />&nbsp;<asp:TextBox runat="server" Width="80px" Font-Size="XX-Small"
                            Font-Names="Verdana" ID="txtSearchCriteriaAllCustomers"></asp:TextBox>
                    <asp:Button ID="btnGo" runat="server" Text="go" OnClick="btnGo_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbIncludeSuspendedUsers" runat="server"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="+ suspended users" AutoPostBack="true"
                        OnCheckedChanged="cbIncludeSuspendedUsers_CheckedChanged" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnAddNewUser" runat="server"
                        Text="add new user" OnClick="btnAddNewUser_Click" />
                    &nbsp;&nbsp;<asp:Button ID="btnExportUserList" runat="server" Text="export user list"
                        OnClick="btnExportUserDetails_click" />
                    <asp:Button ID="btnUserGroups" runat="server" OnClick="btnUserGroups_Click" Text="groups" />
                </td>
                <td align="right" valign="middle" style="white-space: nowrap; height: 27px;">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="2" style="white-space: nowrap">
                    <asp:Label ID="Label1" runat="server" Font-Size="X-Small" Font-Names="Verdana">Selected Customer: </asp:Label>
                    <asp:Label runat="server" ForeColor="Gray" ID="lblCustomerName" Font-Size="X-Small"
                        Font-Names="Verdana"></asp:Label>
                </td>
            </tr>
        </table>
        <br />
        <asp:DataGrid ID="dgSystemAdministrator" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" OnPageIndexChanged="dgSystemAdministrator_Page_Change" AllowPaging="True"
            Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True"
            OnItemCommand="Edit_User">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <PagerStyle Font-Size="X-Small" Font-Names="Verdana" Font-Bold="True" HorizontalAlign="Center"
                ForeColor="Blue" BackColor="Silver" Wrap="False" Mode="NumericPages"></PagerStyle>
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <ItemStyle BackColor="LightGray"></ItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key">
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn Visible="False" DataField="CustomerKey"></asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <asp:Button ID="btnEdit" CommandName="Properties" runat="server" Text="edit" />
                        <asp:Button ID="btnEmail" CommandName="Email" runat="server" Text="email" />
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:BoundColumn DataField="UserId" SortExpression="UserId" HeaderText="User Id">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="UserName" SortExpression="UserName" HeaderText="User Name">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Type" SortExpression="Type" HeaderText="Privilege">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="CustomerAccountCode" SortExpression="CustomerAccountCode"
                    HeaderText="Customer">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Status" SortExpression="Status" HeaderText="Status">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="EmailAddr" SortExpression="EmailAddr" HeaderText="Email Address">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    <ItemTemplate>
                        <asp:Button ID="btnProductProfile" Visible='<%# bSetProductProfileButtonVisibility(Container.DataItem) %>'
                            CommandName="ProductProfile" runat="server" Text="product profile" />
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid>
        <div id="divSystemUserControls" runat="server" visible="false">
            <asp:Label ID="Label17" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Display</asp:Label>
            <asp:DropDownList ID="ddlUsersPerSystemUserPage" runat="server" AutoPostBack="True"
                Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlUsersPerSystemUserPage_SelectedIndexChanged">
                <asp:ListItem Selected="True">10</asp:ListItem>
                <asp:ListItem>50</asp:ListItem>
                <asp:ListItem>250</asp:ListItem>
            </asp:DropDownList>
            <asp:Label ID="Label24" runat="server" Font-Size="XX-Small" Font-Names="Verdana">users/page</asp:Label>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label25" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Order by:</asp:Label>
            <asp:RadioButton ID="rbSystemUserUserID" runat="server" AutoPostBack="True" Checked="True"
                Font-Names="Verdana" Font-Size="XX-Small" GroupName="OrderBySystemUser" Text="User ID"
                OnCheckedChanged="rbSystemUserUserID_CheckedChanged" />
            <asp:RadioButton ID="rbSystemUserUserName" runat="server" AutoPostBack="True" Font-Names="Verdana"
                Font-Size="XX-Small" GroupName="OrderBySystemUser" Text="User Name" OnCheckedChanged="rbSystemUserUserName_CheckedChanged" />
            <asp:RadioButton ID="rbSystemUserCustomerName" runat="server" AutoPostBack="True"
                Font-Names="Verdana" Font-Size="XX-Small" GroupName="OrderBySystemUser" Text="Customer Name"
                OnCheckedChanged="rbSystemUserCustomerName_CheckedChanged" />
        </div>
    </asp:Panel>
    <br />
    <asp:Label ID="lblSystemUserMessage" runat="server" ForeColor="Gray" Font-Size="X-Small"
        Font-Names="Verdana"></asp:Label>
    <asp:Panel ID="pnlCustomerUser" runat="server" Visible="False" Width="100%">
        <table style="width: 100%; font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td align="left" valign="middle" style="white-space: nowrap">
                    <asp:Button ID="btnShowAllCustomerProfiles" OnClick="btnShowAllCustomerProfiles_click"
                        runat="server" Text="show all users" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label14" Font-Size="XX-Small" Font-Names="Verdana"
                        runat="server" Text="search" />&nbsp;<asp:TextBox runat="server" Width="80px" Font-Size="XX-Small"
                            Font-Names="Arial" ID="txtSearchCriteriaCustomer"></asp:TextBox>
                    <asp:Button ID="btnSSearchUsers" OnClick="btn_SearchUsers_Click" runat="server" Text="go" />&nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox
                        ID="cbIncludeSuspendedUsers2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="+ suspended users" AutoPostBack="true" OnCheckedChanged="cbIncludeSuspendedUsers_CheckedChanged" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnSystemUserAddNewUser" OnClick="btn_AddUser_click"
                        runat="server" Text="add new user" />
                    &nbsp;&nbsp;<asp:Button ID="Button20" runat="server" Text="export user list" OnClick="btnExportUserDetails_click" />
                    <asp:Button ID="Button3" runat="server" OnClick="btnUserGroups_Click" Text="groups" />
                </td>
                <td align="right" valign="middle" style="white-space: nowrap">
                    &nbsp;
                </td>
            </tr>
        </table>
        <br />
        <asp:DataGrid ID="dgSuperUser" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"
            OnPageIndexChanged="dgSuperUser_Page_Change" AllowPaging="True" Visible="False"
            AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="Edit_User"
            PageSize="11" OnItemDataBound="dgSuperUser_ItemDataBound">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <PagerStyle Font-Size="X-Small" Font-Bold="True" HorizontalAlign="Center" ForeColor="Blue"
                BackColor="Silver" Wrap="False" Mode="NumericPages"></PagerStyle>
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <ItemStyle BackColor="LightGray"></ItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key">
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn Visible="False" DataField="CustomerKey"></asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <asp:Button ID="btnSuperUserProperties" CommandName="Properties" runat="server" Text="edit" />
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:BoundColumn DataField="UserId" SortExpression="UserId" HeaderText="User Id">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="UserName" SortExpression="UserName" HeaderText="User Name">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Type" SortExpression="Type" HeaderText="Privilege">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Department" SortExpression="Department" HeaderText="Cost Centre">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Status" SortExpression="Status" HeaderText="Status">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="EmailAddr" SortExpression="EmailAddr" HeaderText="Email Address">
                    <HeaderStyle Font-Bold="True" ForeColor="Gray"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    <ItemTemplate>
                        <asp:Button ID="btnSuperUserProductProfile" Visible='<%# bSetProductProfileButtonVisibility(Container.DataItem) %>'
                            CommandName="ProductProfile" runat="server" Text="product profile" />
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid>
        <div id="divSuperUserControls" runat="server" visible="false">
            <asp:Label ID="lbl00x47" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Display</asp:Label>
            <asp:DropDownList ID="ddlUserPerSuperUserPage" runat="server" AutoPostBack="True"
                Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlUserPerSuperUserPage_SelectedIndexChanged">
                <asp:ListItem Selected="True">10</asp:ListItem>
                <asp:ListItem>50</asp:ListItem>
                <asp:ListItem>250</asp:ListItem>
            </asp:DropDownList>
            <asp:Label ID="Label9" runat="server" Font-Size="XX-Small" Font-Names="Verdana">users/page</asp:Label>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label26" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Order by:</asp:Label>
            <asp:RadioButton ID="rbSuperUserUserID" runat="server" AutoPostBack="True" Checked="True"
                Font-Names="Verdana" Font-Size="XX-Small" GroupName="OrderBySuperUser" Text="User ID"
                OnCheckedChanged="rbSuperUserUserID_CheckedChanged" />
            <asp:RadioButton ID="rbSuperUserUserName" runat="server" AutoPostBack="True" Font-Names="Verdana"
                Font-Size="XX-Small" GroupName="OrderBySuperUser" Text="User Name" OnCheckedChanged="rbSuperUserUserName_CheckedChanged" />
        </div>
        <br />
        <asp:Label ID="lblSuperUserMessage" runat="server" ForeColor="Gray" Font-Size="X-Small"
            Font-Names="Verdana"></asp:Label>
    </asp:Panel>
    <asp:Panel ID="pnlAddEditUser" runat="server" Visible="False" Width="100%">
        <table style="font-family: Verdana; font-size: xx-small; color: Gray" width="100%">
            <tr valign="middle">
                <td align="left" valign="middle">
                    <asp:Label ID="lbl001" runat="server" ForeColor="Gray" Font-Size="X-Small" Font-Names="Verdana"
                        Font-Bold="True" Text="User Profile"></asp:Label>
                </td>
                <td align="Right" valign="middle">
                    <asp:Button ID="Button16" runat="server" Text="help" Visible="False" OnClientClick="javascript:OpenHelpWindow('./help/usermngr_hlp.aspx');" />
                    &nbsp;&nbsp;<asp:Button ID="btnReturnToMyPanel" OnClick="btn_ReturnToMyPanel_click"
                        runat="server" Text="go back" />
                </td>
            </tr>
            <tr valign="middle">
                <td valign="middle" colspan="2" align="left">
                    <asp:Label ID="lbl002" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Use this page to add or edit a user's profile. The settings on this page control what the user can view and what email alerts will be sent.</asp:Label>
                    <br />
                </td>
            </tr>
        </table>
        <table id="tbl001" width="100%" style="font-family: Verdana; font-size: xx-small;
            color: Gray">
            <tr>
                <td valign="top">
                    <table style="color: Gray; width: 95%; font-size: xx-small; font-family: Verdana">
                        <tr>
                            <td style="width: 20px">
                            </td>
                            <td style="width: 300px">
                                <asp:Label ID="Label4" runat="server" Font-Bold="True" Text="Website Usage<br />" />
                            </td>
                            <td style="width: 300px" />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label5" runat="server">Book Courier Collections</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkAbleToCreateCollectionRequest" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label15" runat="server">View (but not order) Stock</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkAbleToViewStock" Font-Size="XX-Small" AutoPostBack="True"
                                    OnCheckedChanged="chkAbleToViewStock_CheckedChanged" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:Label ID="Label6" runat="server">Order Stock</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkAbleToCreateStockBookings" Font-Size="XX-Small"
                                    AutoPostBack="True" OnCheckedChanged="chkAbleToCreateStockBookings_CheckedChanged" />
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 18px" />
                            <td style="height: 18px">
                                <asp:Label ID="Label7" runat="server" Font-Bold="True" Text="Stock Emailing Options<br />" />
                            </td>
                            <td style="height: 18px" />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label8" runat="server" Text="Receive Own Booking Confirmation" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkStockBookingAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblLegendReceiveAllBookingConfirmations" runat="server">Receive ALL Booking Confirmations</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkStockBookingAlertAll" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label10" runat="server" Text="Receive Goods-In Alerts" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkStockArrivalAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label11" runat="server" Text="Receive Low Stock Alerts" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkLowStockAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trCourierEmailingOptions01" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="Label12" runat="server" Font-Bold="True" Text="Courier Emailing Options<br />" />
                            </td>
                            <td />
                        </tr>
                        <tr id="trCourierEmailingOptions02" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="Label5a" runat="server">Receive Own Booking Confirmation</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkCourierBookingAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trCourierEmailingOptions03" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="Label5b" runat="server" Text="Receive ALL Booking Confirmations" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkCourierBookingAlertAll" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trCourierEmailingOptions04" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="Label5c" runat="server" Text="Receive AWB Despatch Alerts" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkAWBDespatchAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trCourierEmailingOptions05" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="Label5d" runat="server" Text="Receive AWB Delivery Confirmation" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkAWBDeliveryAlert" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trUserPublicationOptions01" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="Label12upo" runat="server" Font-Bold="True" Text="User Publication Options<br />" />
                            </td>
                            <td />
                        </tr>
                        <tr id="trUserPublicationOptions02" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="Label5upo2" runat="server">Receive Own User Publications Booking Confirmations</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="cbUserPublicationsOwnBookingAlerts" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trUserPublicationOptions03" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="Label5upo2a" runat="server">Receive ALL User Publications Booking Confirmations</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="cbUserPublicationsAllBookingAlerts" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trUserPublicationOptions04" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="Label5bupo3" runat="server" Text="Receive Own Publication Inactivity Alerts" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="cbUserPublicationsOwnInactivityAlerts" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr id="trUserPublicationOptions05" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="Label5bupo3a" runat="server" Text="Receive ALL Publication Inactivity Alerts" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="cbUserPublicationsAllInactivityAlerts" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label5e" runat="server" Font-Bold="True" Text="Address Book" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Labelf" runat="server">View Global Address Book</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkViewGlobalAddressBook" Font-Size="XX-Small" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label5g" runat="server" Text="Edit Global Address Book" />
                            </td>
                            <td>
                                <asp:CheckBox runat="server" ID="chkEditGlobalAddressBook" Font-Size="XX-Small" />
                            </td>
                        </tr>
                    </table>
                </td>
                <td valign="top">
                    <table style="color: Gray; font-size: xx-small; font-family: Verdana">
                        <tr>
                            <td style="width: 20px">
                            </td>
                            <td style="width: 150px">
                                <asp:Label ID="lblCompany989" runat="server">Company</asp:Label>
                            </td>
                            <td style="width: 210px">
                                <asp:Label runat="server" ID="lblCustomer" ForeColor="Navy"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label runat="server" ID="lblFirstName" ForeColor="Red">First Name</asp:Label>
                                &nbsp;
                                <asp:RequiredFieldValidator ID="rfvFirstName" runat="server" ControlToValidate="txtFirstName">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="1" Font-Names="Verdana" ID="txtFirstName"
                                    MaxLength="50" Font-Size="XX-Small"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label runat="server" ID="lblLastName" ForeColor="Red">Last Name</asp:Label>
                                &nbsp;
                                <asp:RequiredFieldValidator ID="rfvLastName" runat="server" ControlToValidate="txtLastName">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="2" Font-Names="Verdana" ID="txtLastName"
                                    MaxLength="50" Font-Size="XX-Small"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label runat="server" ID="lblUserId" ForeColor="Red">User ID</asp:Label>
                                &nbsp;
                                <asp:RequiredFieldValidator ID="rfvUserId" runat="server" ControlToValidate="txtUserId">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="3" Font-Names="Verdana" Font-Size="XX-Small"
                                    ID="txtUserId" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label runat="server" ID="lblPassword" ForeColor="Red">Password</asp:Label>
                                &nbsp;
                                <asp:RequiredFieldValidator ID="rfvPassword" runat="server" ControlToValidate="txtPassword">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="4" Font-Names="Verdana" Font-Size="XX-Small"
                                    ID="txtPassword" MaxLength="12"></asp:TextBox>
                                <asp:CheckBox ID="cbForcePasswordChange" Text="force password change" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="Label18" runat="server">User cannot change password</asp:Label>
                            </td>
                            <td>
                                <asp:CheckBox ID="cbUserCannotChangePassword" Text="" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label runat="server" ID="lblEmailAddr" ForeColor="Red">Email Address</asp:Label>
                                &nbsp;
                                <asp:RequiredFieldValidator ID="rfvEmailAddr" runat="server" ControlToValidate="txtEmailAddr">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Font-Names="Verdana" Width="180px"
                                    Font-Size="XX-Small" ID="txtEmailAddr" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trTelephone" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="lblTelephoneNo" runat="server">Telephone No</asp:Label>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Font-Names="Verdana" Width="180px"
                                    Font-Size="XX-Small" ID="txtTelephone" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trHysterDealershipCode" runat="server">
                            <td />
                            <td>
                                <asp:Label ID="lblDealerShipCode" runat="server">Dealership Code / Nacco Location Code</asp:Label>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Font-Names="Verdana" Width="180px"
                                    Font-Size="XX-Small" ID="txtDealershipCode" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblCollectionPoint" runat="server" Text="Collection Point" />
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Font-Names="Verdana" Width="180px"
                                    Font-Size="XX-Small" ID="txtCollectionPoint" MaxLength="100"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblAccessLevel" runat="server">Access Level</asp:Label>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="6" Font-Names="Verdana" Font-Size="XX-Small"
                                    Enabled="False" ID="txtAccessLevel" MaxLength="20"></asp:TextBox>&nbsp;&nbsp;<asp:LinkButton
                                        ID="lnkbtnPromoteToSuperUser" OnClick="lnkbtnPromoteToSuperUser_Click" OnClientClick='return confirm("Are you sure you want to promote this user to superuser privilege?");'
                                        runat="server" Width="70px" Style="height: 12px">promote&nbsp;to&nbsp;superuser</asp:LinkButton>&nbsp;<a
                                            id="aAccessLevelHelp" visible="true" runat="server" onmouseover="return escape('Gives the selected user <b>superuser</b> privilege. Once the user has been granted this access level, it cannot be revoked.')"
                                            style="color: gray; cursor: help">&nbsp;&nbsp;&nbsp;&nbsp;?&nbsp;</a>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblDepartment" runat="server" Text="Cost Centre" />
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="7" Font-Names="Verdana" Font-Size="XX-Small"
                                    ID="txtDepartment" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trUserGroup" runat="server" visible="false">
                            <td>
                            </td>
                            <td>
                                <asp:Label ID="lblUserGroup" runat="server" ForeColor="Red">User Group</asp:Label>&nbsp;
                                &nbsp;<asp:RequiredFieldValidator ID="rfvUserGroup" runat="server" ControlToValidate="ddlUserGroup"
                                    EnableClientScript="False" InitialValue="0">###</asp:RequiredFieldValidator>
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlUserGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblUserStatus" runat="server">User Status</asp:Label>
                            </td>
                            <td>
                                <asp:RadioButtonList runat="server" ForeColor="Gray" TabIndex="8" Font-Size="XX-Small"
                                    Font-Names="Verdana" RepeatDirection="Horizontal" ID="btnlst_UserStatus">
                                    <asp:ListItem Value="Active">Active</asp:ListItem>
                                    <asp:ListItem Value="Suspended">Suspended</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:Label ID="lblInternalUser" runat="server" Font-Bold="True" Font-Names="Verdana"
                                    Font-Size="XX-Small" ForeColor="Red" Text="internal"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <asp:Label ID="lblRunningHeader" runat="server">Running Header</asp:Label>
                            </td>
                            <td>
                                <asp:TextBox runat="server" ForeColor="Navy" TabIndex="9" Font-Names="Verdana" Width="180px"
                                    Font-Size="XX-Small" ID="txtRunningHeaderImage" MaxLength="100">default</asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trProductCreditsStatus1" runat="server" visible="false">
                            <td />
                            <td>
                                <asp:Label ID="lblLegendProductCreditStatus" runat="server">Product Credit Status</asp:Label>
                            </td>
                            <td>
                                <asp:LinkButton ID="lnkbtnAdjustCredits" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    OnClick="lnkbtnAdjustCredits_Click">adjust&nbsp;product&nbsp;credits</asp:LinkButton>
                            </td>
                        </tr>
                        <tr id="trProductCreditsStatus2" runat="server" visible="false">
                            <td />
                            <td colspan="2">
                                <asp:TextBox ID="tbProductCreditsStatus" runat="server" Rows="4" TextMode="MultiLine"
                                    Width="100%" ReadOnly="True" Font-Names="Arial" Font-Size="XX-Small" Wrap="False"
                                    ToolTip="Product Code | Product Description | Remaining Credit | Credit Expiry"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td />
                            <td align="right" runat="server">
                                <br />
                                <asp:Button ID="Button12" OnClick="btn_SaveUserProfileChanges_click" runat="server"
                                    Text="save" Width="80px" />
                                &nbsp;&nbsp;<asp:Button ID="Button1" OnClick="btn_ReturnToMyPanel_click" runat="server"
                                    Text="cancel" CausesValidation="False" Width="80px" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td valign="top">
                </td>
                <td valign="top">
                </td>
            </tr>
            <tr>
                <td valign="top">
                </td>
                <td valign="top" align="right">
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hidCustomer" runat="server" />
        &nbsp;
    </asp:Panel>
    <asp:Panel ID="pnlEmail" runat="server" Visible="false" Width="100%">
        <table style="font-family: Verdana; font-size: x-small; color: Gray" width="100%">
            <tr valign="middle">
                <td align="left" valign="middle">
                    <asp:Label ID="Label2" runat="server" ForeColor="Gray" Font-Size="X-Small" Font-Names="Verdana"
                        Font-Bold="True">Email User Access Details</asp:Label>
                </td>
                <td align="right" valign="middle">
                    <asp:Button ID="Button17" runat="server" Text="help" Visible="False" OnClientClick="javascript:OpenHelpWindow('./help/profile_hlp.aspx');" />
                    &nbsp;&nbsp;<asp:Button ID="Button2" OnClick="btn_ReturnToMyPanel_click" runat="server"
                        Text="go back" />
                </td>
            </tr>
            <tr valign="middle">
                <td valign="middle" colspan="2" align="left">
                    <asp:Label ID="Label3" runat="server" Font-Size="XX-Small" Font-Names="Verdana">To email account
                        access details to the specified address, modify the message text as required, then click <b>Send Email</b>.</asp:Label>
                    <br />
                </td>
            </tr>
        </table>
        <br />
        <asp:TextBox ID="tbAccessDetailsMessageText" runat="server" Height="64px" TextMode="MultiLine"
            Width="401px"></asp:TextBox><br />
        <br />
        <table style="font-size: x-small; color: gray; font-family: Verdana">
            <tr>
                <td style="width: 120px">
                    User ID:
                </td>
                <td style="width: 17px">
                </td>
                <td style="width: 145px">
                    <asp:Label ID="lblAccessDetailsUserID" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="width: 120px; height: 18px">
                    Password:
                </td>
                <td style="width: 17px; height: 18px">
                </td>
                <td style="width: 145px; height: 18px">
                    <asp:Label ID="lblAccessDetailsPassword" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="width: 120px">
                    Email Address:
                </td>
                <td style="width: 17px">
                </td>
                <td style="width: 145px">
                    <asp:TextBox ID="tbAccessDetailsEmail" runat="server" Width="250px"></asp:TextBox>
                </td>
            </tr>
        </table>
        <br />
        <asp:Button ID="btnSendEmail" runat="server" Text="send email" OnClick="btnSendEmail_Click" /><br />
    </asp:Panel>
    <asp:Panel ID="pnlProductProfile" runat="server" Visible="False" Width="100%">
        <asp:Label ID="Label5z" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="Gray" Text="User Product Profile"></asp:Label>
        <table style="font-family: Verdana; font-size: x-small" width="100%">
            <tr valign="middle">
                <td colspan="2" align="left" valign="middle" style="font-size: xx-small; color: gray;
                    font-family: Verdana">
                    The settings on this page control what products the selected user can view for ordering,
                    and for each product whether there is maximum per-order quantity limit. A maximum
                    order quantity of 0 signifies no limit.<br />
                    <br />
                </td>
            </tr>
            <tr valign="middle">
                <td align="left" valign="middle" style="white-space: nowrap">
                    <asp:Button ID="btnShowAllProducts" OnClick="btn_ShowAllProducts_click" runat="server"
                        Text="show all products..." />
                    &nbsp;<asp:CheckBox ID="cbShowAllowToOrder" runat="server" Checked="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="allowed to order" />
                    &nbsp;<asp:CheckBox ID="cbShowNotAllowToOrder" runat="server" Checked="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="NOT allowed to order" />
                    &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label
                        ID="Label4a" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="search"></asp:Label>
                    &nbsp;<asp:TextBox runat="server" Width="120px" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtUserProfileProdSeach" />
                    <asp:Button ID="btnSearchProducts" OnClick="btn_SearchProducts_click" runat="server"
                        Text="go" />
                </td>
                <td align="right" valign="middle" style="white-space: nowrap">
                    <asp:Button ID="Button13" OnClick="btn_SaveUserProductProfileChanges_click" runat="server"
                        Text="save changes" />
                    <asp:Button ID="Button5" OnClick="btn_ReturnToMyPanel_click" runat="server" Text="go back" />
                </td>
            </tr>
            <tr>
                <td valign="middle" style="white-space: nowrap">
                    <asp:Label ID="lbl004" runat="server">User: </asp:Label>
                    <asp:Label runat="server" ForeColor="Navy" ID="lblProdProfileUserName"></asp:Label>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    &nbsp;&nbsp;
                    <asp:Label ID="lblProductProfileSearchResult" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" ForeColor="Red" Visible="False">No matching records</asp:Label>
                </td>
                <td align="right" valign="middle" style="white-space: nowrap">
                    <asp:Label ID="lblDefaultMaxGrabQty" runat="server" Visible="False" Font-Names="Verdana"
                        Font-Size="XX-Small">Default max order
                        qty:</asp:Label>&nbsp;
                    <asp:TextBox ID="txtDefaultGrabQty" runat="server" Visible="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="50px" Height="20">0</asp:TextBox>&nbsp;
                </td>
            </tr>
        </table>
        <asp:DataGrid ID="dgUserProducts" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" AllowSorting="True" OnSortCommand="SortUsersProductsGrid" AllowPaging="True"
            OnPageIndexChanged="dgUserProducts_PageIndexChanged">
            <HeaderStyle Font-Size="10pt" Font-Names="Arial" Wrap="False" BorderColor="Gray">
            </HeaderStyle>
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <ItemStyle Font-Size="XX-Small" Font-Names="Arial" BackColor="LightGray"></ItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key" HeaderText="Key"></asp:BoundColumn>
                <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Code">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Blue"
                        VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="ProductDate" SortExpression="ProductDate" HeaderText="Date">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription"
                    HeaderText="Description">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                </asp:BoundColumn>
                <asp:BoundColumn Visible="False" DataField="LanguageId" SortExpression="LanguageId"
                    HeaderText="Language">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn Visible="False" DataField="ProductDepartmentId" SortExpression="ProductDepartmentId"
                    HeaderText="Cost Centre">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="ProductCategory" SortExpression="ProductCategory" HeaderText="Category">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="SubCategory" SortExpression="SubCategory" HeaderText="Sub Category">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="Allow Pick">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        HorizontalAlign="Center" ForeColor="Gray" Width="6%" VerticalAlign="Bottom">
                    </HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnTogglePickButton" OnClick="btn_ToggleAllowPick_Click" runat="server"
                            Text="select all" />
                        <br />
                        <asp:Label runat="server" Font-Size="XX-Small" Font-Names="Verdana">allow to order</asp:Label>&nbsp;
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox Checked='<%# DataBinder.Eval(Container, "DataItem.AbleToPick") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Apply Max Grab">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        HorizontalAlign="Center" ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleMaxGrabButton" runat="server" Text="select all" OnClick="btnToggleMaxGrabButton_Click" />
                        <br />
                        <asp:Label runat="server" Font-Size="XX-Small" Font-Names="Verdana">apply max order</asp:Label>&nbsp;
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox Checked='<%# DataBinder.Eval(Container, "DataItem.ApplyMaxGrab ") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Quantity">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        HorizontalAlign="Center" ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnApplyMaxGrabQty" OnClick="btn_ApplyMaxGrabQty_Click" runat="server"
                            Text="apply qty" />
                        <br />
                        <asp:Label ID="Label1" runat="server" Font-Size="XX-Small" Font-Names="Verdana">max
                            order qty</asp:Label>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="txtMaxGrabQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="50px" Text='<%# DataBinder.Eval(Container, "DataItem.MaxGrabQty") %>'></asp:TextBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
            <PagerStyle HorizontalAlign="Center" Mode="NumericPages" />
        </asp:DataGrid>
        <table width="100%" id="tblSaveCancelProductProfile" runat="server" visible="false"
            style="font-family: Verdana; font-size= x-small">
            <tr>
                <td align="left">
                    <asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Products / page:" />
                    &nbsp;<asp:DropDownList ID="ddlUsersPerPage" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnSelectedIndexChanged="ddlUsersPerPage_SelectedIndexChanged">
                        <asp:ListItem Selected="True">10</asp:ListItem>
                        <asp:ListItem>50</asp:ListItem>
                        <asp:ListItem>200</asp:ListItem>
                        <asp:ListItem>1000</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;<asp:Label ID="lblSortValue1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="(Sorting on "></asp:Label>
                    <asp:Label ID="lblSortValue" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Product Code"></asp:Label><asp:Label ID="lblSortValue0"
                            runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=")" />
                </td>
                <td align="right" valign="middle">
                    <asp:Button ID="Button13bis" OnClick="btn_SaveUserProductProfileChanges_click" runat="server"
                        Text="save changes" />
                    &nbsp;&nbsp;<asp:Button ID="Button6" OnClick="btn_ReturnToMyPanel_click" runat="server"
                        Text="cancel" />
                </td>
            </tr>
        </table>
        <asp:Label ID="lblProductProfileMessage" runat="server" ForeColor="#00C000" Font-Size="X-Small"></asp:Label>
    </asp:Panel>
    <asp:Panel ID="pnlChooseUserType" runat="server" Visible="False" Width="100%">
        <div align="center">
            <table width="700px" style="font-family: Verdana; font-size: x-small">
                <tr align="center">
                    <td align="center">
                        <br />
                        <br />
                        <br />
                        <asp:Label ID="lbl005" runat="server" ForeColor="Gray" Font-Size="Small" Font-Names="Verdana">Select type of user to create</asp:Label>
                        <br />
                        <br />
                    </td>
                </tr>
                <tr align="center">
                    <td align="center">
                        <asp:Panel ID="pnlsaUser" runat="server">
                            <asp:RadioButtonList ID="rblSaUser" OnSelectedIndexChanged="rblSaUser_IndexChanged"
                                AutoPostBack="True" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                                ForeColor="Gray">
                                <asp:ListItem Value="AccountHandler">Account Handler</asp:ListItem>
                                <asp:ListItem Value="SuperUser">Super User</asp:ListItem>
                                <asp:ListItem Value="Product Owner">Product Owner</asp:ListItem>
                                <asp:ListItem Value="User">User</asp:ListItem>
                            </asp:RadioButtonList>
                        </asp:Panel>
                        <asp:Panel ID="pnlInternalSuperUserEnhanced" runat="server">
                            <asp:RadioButtonList ID="rblInternalSuperUserWithProductOwner" AutoPostBack="True"
                                runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Gray" OnSelectedIndexChanged="rblInternalSuperUserWithProductOwner_SelectedIndexChanged">
                                <asp:ListItem Value="AccountHandler">Account Handler</asp:ListItem>
                                <asp:ListItem Value="SuperUser">Super User</asp:ListItem>
                                <asp:ListItem Value="Product Owner">Product Owner</asp:ListItem>
                                <asp:ListItem Value="User">User</asp:ListItem>
                            </asp:RadioButtonList>
                        </asp:Panel>
                        <asp:Panel ID="pnlInternalSuperUser" runat="server">
                            <asp:RadioButtonList ID="rblInternalSuperUser" OnSelectedIndexChanged="rblInternalSuperUser_IndexChanged"
                                AutoPostBack="True" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                                ForeColor="Gray">
                                <asp:ListItem Value="AccountHandler">Account Handler</asp:ListItem>
                                <asp:ListItem Value="SuperUser">Super User</asp:ListItem>
                                <asp:ListItem Value="User">User</asp:ListItem>
                            </asp:RadioButtonList>
                        </asp:Panel>
                        <asp:Panel ID="pnlSuperUserWithProductOwner" runat="server">
                            <asp:RadioButtonList ID="rblSuperUserWithProductOwner" AutoPostBack="True" runat="server"
                                Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Gray" OnSelectedIndexChanged="rblSuperUserWithProductOwner_SelectedIndexChanged">
                                <asp:ListItem Value="SuperUser">Super User</asp:ListItem>
                                <asp:ListItem Value="Product Owner">Product Owner</asp:ListItem>
                                <asp:ListItem Value="User">User</asp:ListItem>
                            </asp:RadioButtonList>
                        </asp:Panel>
                        <asp:Panel ID="pnlSuperUser" runat="server">
                            <asp:RadioButtonList ID="rblSuperUser" OnSelectedIndexChanged="rblSuperUser_IndexChanged"
                                AutoPostBack="True" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                                ForeColor="Gray">
                                <asp:ListItem Value="SuperUser">Super User</asp:ListItem>
                                <asp:ListItem Value="User">User</asp:ListItem>
                            </asp:RadioButtonList>
                        </asp:Panel>
                    </td>
                </tr>
                <tr align="center">
                    <td align="center">
                        <br />
                        &nbsp;&nbsp;
                        <asp:Button ID="Button15" OnClick="btn_ContinueToAddUser_click" runat="server" Text="continue" />
                        <asp:Button ID="Button8" OnClick="btn_ReturnToMyPanel_click" runat="server" Text="cancel" />
                    </td>
                </tr>
            </table>
        </div>
    </asp:Panel>
    <asp:Panel ID="pnlDatabaseError" runat="server" Visible="False" Width="100%">
        <p>
            <asp:Table ID="Table15" runat="server" Width="100%" Font-Names="Arial" Font-Size="X-Small">
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell HorizontalAlign="Center">
                            <asp:Image runat="server" ImageUrl="./images/icon_shutdown.gif" ></asp:Image>
                            &nbsp;<asp:Label runat="server" forecolor="Blue" font-size="Small" font-names="Arial" Text="An error has occurred" />
                    </asp:TableCell></asp:TableRow><asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell HorizontalAlign="Center">
                            <asp:Label runat="server" font-size="X-Small" font-names="Arial" Text="Please contact customer
                            services for assistance or" /> &nbsp;<asp:LinkButton runat="server" ForeColor="Blue" onclick="btn_ReturnToStart">click here</asp:LinkButton>
                            &nbsp;<asp:Label runat="server" font-size="X-Small" font-names="Arial" Text="continue" />
                    </asp:TableCell></asp:TableRow></asp:Table><asp:Table ID="Table16" runat="server" Width="100%" Font-Names="Arial" Font-Size="X-Small">
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell HorizontalAlign="Center">
                        <asp:Label ID="lblDBError" runat="server" Font-Size="X-Small" Font-Names="Arial"
                            ForeColor="Red"></asp:Label>
                    </asp:TableCell></asp:TableRow></asp:Table></p></asp:Panel><asp:Panel ID="pnlShowNoProductProfileMessage" runat="server" Visible="False" Width="100%">
        <asp:Table ID="tbl006" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
            <asp:TableRow VerticalAlign="Middle">
                <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" Wrap="False"></asp:TableCell><asp:TableCell
                    HorizontalAlign="Right" VerticalAlign="Middle" Wrap="False">
                    <asp:Button ID="Button19" runat="server" Text="Help" Visible="False" OnClientClick="javascript:OpenHelpWindow('./help/consignment_hlp.aspx');" />
                    &nbsp;&nbsp;<asp:Button ID="Button9" OnClick="btn_ReturnToMyPanel_click" runat="server"
                        Text="go back" />&nbsp;
                </asp:TableCell></asp:TableRow></asp:Table><asp:Table ID="tabConfirmDeleteCollection" runat="server" Width="100%" Font-Size="X-Small"
            Font-Names="Verdana" ForeColor="Gray">
            <asp:TableRow>
                <asp:TableCell Wrap="False" HorizontalAlign="Center">
                        <br />
                        <br />
                        <asp:Label runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Super Users don't have Product Profiles (they can view all products)</asp:Label>
                        <br />
                        <br />
                        <br />
                        <br />
                        <br />
                        <br />
                </asp:TableCell></asp:TableRow><asp:TableRow>
                <asp:TableCell HorizontalAlign="Center">
                    <br />
                    <br />
                    <asp:Button ID="Button10" OnClick="btn_ReturnToMyPanel_click" runat="server" Text="continue" />
                    <br />
                    <br />
                </asp:TableCell></asp:TableRow></asp:Table></asp:Panel><asp:Panel ID="pnlConfirmPermissionsChange" runat="server" Visible="False" Width="100%">
        <asp:Table ID="tbl006zzz" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
            <asp:TableRow VerticalAlign="Middle">
                <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" Wrap="False"></asp:TableCell><asp:TableCell
                    HorizontalAlign="Right" VerticalAlign="Middle" Wrap="False">
                    &nbsp;&nbsp;<asp:Button ID="Button9zzz" OnClick="btn_ReturnToMyPanel_click" runat="server"
                        Text="go back" />&nbsp;
                </asp:TableCell></asp:TableRow></asp:Table><asp:Table ID="tabConfirmPermissionsChange" runat="server" Width="100%" Font-Size="X-Small"
            Font-Names="Verdana" ForeColor="Gray">
            <asp:TableRow>
                <asp:TableCell Wrap="False" HorizontalAlign="Center">
                    <br />
                    <br />
                    <asp:Label ID="Label16" runat="server" ForeColor="Gray" Font-Size="X-Small" Font-Names="Verdana"
                        Font-Bold="True">You have changed the default permissions for this user (or no user permissions were previously defined).<br /><br />Click continue to apply the new permissions, or cancel to retain any existing permissions.</asp:Label>
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                    <br />
                </asp:TableCell></asp:TableRow><asp:TableRow>
                <asp:TableCell HorizontalAlign="Center">
                    <br />
                    <br />
                    <asp:Button ID="btnContinueToChangePermissions" runat="server" Text="continue" OnClick="btnContinueToChangePermissions_Click" />
                    &nbsp;&nbsp;
                    <asp:Button ID="Buttona3" OnClick="btn_ReturnToMyPanel_click" runat="server" Text="cancel" />
                    <br />
                    <br />
                </asp:TableCell></asp:TableRow></asp:Table></asp:Panel><asp:Table ID="Table9" runat="server" Width="100%">
        <asp:TableRow>
            <asp:TableCell HorizontalAlign="Right">
                <asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Size="X-Small" Font-Names="Verdana"></asp:Label>
            </asp:TableCell></asp:TableRow></asp:Table><asp:Panel ID="pnlUserGroups" runat="server" Width="100%" Visible="False">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%">
                    <asp:Label ID="Label3xx" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="User groups"></asp:Label></td><td style="width: 29%">
                </td>
                <td style="width: 20%">
                </td>
                <td align="right" style="width: 29%">
                    <asp:Button ID="btnUsersReport" runat="server" OnClick="btnUsersReport_Click" Text="users report" />&nbsp; <asp:Button ID="btnBackFromUserGroups" runat="server" OnClick="btnBackFromUserGroups_Click"
                        Text="back" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%" valign="top">
                    <asp:Label ID="Label1xx" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="User groups:"></asp:Label></td><td colspan="2">
                    <asp:ListBox ID="lbDefinedUserGroups" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Rows="10" Width="100%" OnSelectedIndexChanged="lbDefinedUserGroups_SelectedIndexChanged"
                        AutoPostBack="True"></asp:ListBox>
                </td>
                <td align="left" style="width: 29%" valign="top">
                    &nbsp; &nbsp; <asp:Button ID="btnNewUserGroup" runat="server" Text="new user group" Width="200px"
                        OnClick="btnNewUserGroup_Click" /><br />
                    <br />
                    &nbsp; &nbsp; <asp:Button ID="btnRenameUserGroup" runat="server" Text="rename user group" Width="200px"
                        OnClick="btnRenameUserGroup_Click" Enabled="False" /><br />
                    <br />
                    &nbsp; &nbsp; <asp:Button ID="btnRemoveUserGroup" runat="server" Text="remove user group" Width="200px"
                        OnClick="btnRemoveUserGroup_Click" OnClientClick='return confirm("Are you sure you want to remove this user group? User accounts will NOT be removed. No user must be referencing the group to be deleted.");'
                        Enabled="False" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" align="right">
                </td>
                <td colspan="3">
                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                    &nbsp;&nbsp; </td><td style="width: 1%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlNewUserGroup" runat="server" Width="100%" Visible="False">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%">
                    <asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="New user group"></asp:Label></td><td style="width: 29%">
                </td>
                <td style="width: 20%">
                </td>
                <td align="right" style="width: 29%">
                    <asp:Button ID="btnBackFromNewUserGroup" runat="server" OnClick="btnBackFromNewUserGroup_Click"
                        Text="back" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" align="right">
                    <asp:Label ID="Label20" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="User group name:"></asp:Label></td><td colspan="3">
                    <asp:TextBox ID="tbNewUserGroupName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        MaxLength="50" Width="200px" />
                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                    &nbsp;&nbsp; </td><td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%">
                </td>
                <td colspan="3">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%">
                </td>
                <td colspan="3">
                    <asp:Button ID="btnCreateNewUserGroup" runat="server" Text="create" OnClick="btnCreateNewUserGroup_Click" />
                    &nbsp; &nbsp;<asp:Button ID="btnCancelNewUserGroup" runat="server" Text="cancel"
                        OnClick="btnCancelNewUserGroup_Click" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlRenameUserGroup" runat="server" Width="100%" Visible="False">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%">
                    <asp:Label ID="Label21" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Rename user group"></asp:Label></td><td style="width: 29%">
                </td>
                <td style="width: 20%">
                </td>
                <td align="right" style="width: 29%">
                    <asp:Button ID="btnBackFromRenameUserGroup" runat="server" Text="back" OnClick="btnBackFromRenameUserGroup_Click" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" align="right">
                    <asp:Label ID="Label22" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="User group name:"></asp:Label></td><td colspan="3">
                    <asp:TextBox ID="tbRenameUserGroupNewName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        MaxLength="50" Width="200px" />
                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                    &nbsp;&nbsp; </td><td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%">
                </td>
                <td colspan="3">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%">
                </td>
                <td colspan="3">
                    <asp:Button ID="btnDoRenameUserGroup" runat="server" Text="rename" OnClick="btnDoRenameUserGroup_Click" />
                    &nbsp;&nbsp; <asp:Button ID="btnCancelRenameUserGroup" runat="server" Text="cancel" OnClick="btnCancelRenameUserGroup_Click" />
                </td>
                <td style="width: 1%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlAdjustProductCredits" runat="server" Width="100%" Visible="False">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%">
                    <asp:Label ID="Label3xxapc" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Adjust Product Credits for user "></asp:Label>&nbsp;<asp:Label 
                        ID="lblAdjustProductCreditsUser" runat="server" Font-Bold="True" 
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></td><td style="width: 29%">
                <asp:Button 
                        ID="btnNewProductCredit" runat="server" onclick="btnNewProductCredit_Click" 
                        Text="new product credit" /></td>
                <td style="width: 20%">
                </td>
                <td align="right" style="width: 29%">
                    &nbsp; <asp:Button ID="btnBackFromAdjustProductCredits" runat="server" Text="back" onclick="btnBackFromAdjustProductCredits_Click" /></td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td align="right" style="width: 20%" valign="middle"><asp:Label ID="Label1xxapc" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Product Credits:"></asp:Label></td><td colspan="3">
                    <asp:GridView 
                        ID="gvAdjustProductCredits" runat="server" CellPadding="2" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" ><Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnEditProductCredit" runat="server" OnClick="lnkbtnEditProductCredit_Click"
                                        Font-Names="Verdana" CommandArgument='<%# Container.DataItem("id")%>' Font-Size="XX-Small">edit</asp:LinkButton>&nbsp;<asp:LinkButton
                                            ID="lnkbtnRemoveProductCredit" runat="server" OnClick="lnkbtnRemoveProductCredit_Click"
                                            Font-Names="Verdana" CommandArgument='<%# Container.DataItem("id")%>' OnClientClick="return confirm(&quot;This will remove this credit from the user. Are you sure you want to do this?&quot;);"
                                            Font-Size="XX-Small">remove</asp:LinkButton></ItemTemplate></asp:TemplateField><asp:BoundField 
                                DataField="Product" HeaderText="Product" ReadOnly="True" 
                                SortExpression="Product" /><asp:BoundField DataField="ProductDescription" 
                                HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                                <asp:BoundField 
                                DataField="RemainingCredit" HeaderText="Remaining Credit" ReadOnly="True" 
                                SortExpression="Credit" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField><asp:TemplateField HeaderText="Type" SortExpression="EnforceCreditLimit">
                                <ItemTemplate>
                                    <asp:Label ID="lblProductCreditType" runat="server" Text='<%# gvProductCreditType(Container.DataItem) %>' />
                                </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" /></asp:TemplateField>
                                <asp:BoundField 
                                DataField="CreditStartDateTime" HeaderText="Start" ReadOnly="True" 
                                SortExpression="CreditStartDateTime" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField><asp:BoundField 
                                DataField="CreditEndDateTime" HeaderText="End" ReadOnly="True" 
                                SortExpression="CreditEndDateTime" ><ItemStyle HorizontalAlign="Center" /></asp:BoundField></Columns></asp:GridView></td><td style="width: 1%">
                </td>
            </tr>
            <tr id="trEditProductCredit" runat="server" visible="false">
                <td style="width: 1%">
                </td>
                <td style="width: 20%" align="right">
                </td>
                <td colspan="3">
                <asp:Label ID="lblEditCreditProduct" runat="server" 
                        Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="PRODUCT"></asp:Label>&nbsp;<asp:Label 
                        ID="Label27" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Credit:"></asp:Label>&nbsp;<telerik:RadNumericTextBox 
                        ID="rntbEditCredit" runat="server" DataType="System.Int32" Font-Names="Verdana" 
                        Font-Size="XX-Small" MaxValue="999" MinValue="0" ShowSpinButtons="True" 
                        Width="60px"><NumberFormat DecimalDigits="0" ZeroPattern="n" /></telerik:RadNumericTextBox>&nbsp;&nbsp;<asp:CheckBox 
                        ID="cbEditCreditEnforce" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Enforce" TextAlign="Left" />&nbsp;<asp:Label 
                        ID="Label28" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Start:"></asp:Label>&nbsp;<telerik:RadDateTimePicker 
                        ID="rdtpEditCreditStartDateTime" runat="server" Culture="en-GB" 
                        DateInput-DateFormat="dd-MMM-yyyy " DateInput-DisplayDateFormat="dd-MMM-yyyy " 
                        FocusedDate="2013-01-01" Font-Names="Verdana" Font-Size="XX-Small" 
                        MinDate="2013-01-01" Width="180px"><TimeView CellSpacing="-1" Culture="en-GB"></TimeView><TimePopupButton 
                            HoverImageUrl="" ImageUrl="" /><Calendar 
                            UseColumnHeadersAsSelectors="False" UseRowHeadersAsSelectors="False" 
                            ViewSelectorText="x"></Calendar><DateInput DateFormat="dd-MMM-yyyy hh:mm" 
                            DisplayDateFormat="dd-MMM-yyyy hh:mm" LabelWidth="40%"></DateInput><DatePopupButton 
                            HoverImageUrl="" ImageUrl="" /></telerik:RadDateTimePicker>&nbsp; <asp:Label 
                        ID="Label29" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="End:"></asp:Label>&nbsp;<telerik:RadDateTimePicker 
                        ID="rdtpEditCreditEndDateTime" runat="server" Culture="en-GB" 
                        DateInput-DateFormat="dd-MMM-yyyy " DateInput-DisplayDateFormat="dd-MMM-yyyy " 
                        FocusedDate="2013-01-01" Font-Names="Verdana" Font-Size="XX-Small" 
                        MinDate="2013-01-01" Width="180px"><TimeView CellSpacing="-1" Culture="en-GB"></TimeView><TimePopupButton 
                            HoverImageUrl="" ImageUrl="" /><Calendar 
                            UseColumnHeadersAsSelectors="False" UseRowHeadersAsSelectors="False" 
                            ViewSelectorText="x"></Calendar><DateInput DateFormat="dd-MMM-yyyy hh:mm" 
                            DisplayDateFormat="dd-MMM-yyyy hh:mm" LabelWidth="40%"></DateInput><DatePopupButton 
                            HoverImageUrl="" ImageUrl="" /></telerik:RadDateTimePicker>&nbsp;<asp:Button 
                        ID="btnEditCreditSave" runat="server" Text="save" 
                        onclick="btnEditCreditSave_Click" />&nbsp;<asp:Button 
                        ID="btnEditCreditCancel" runat="server" Text="cancel" 
                        onclick="btnEditCreditCancel_Click" /></td><td style="width: 1%">
                </td>
            </tr>
            <tr id="trNewProductCredit" runat="server" visible="false">
                <td style="width: 1%">
                </td>
                <td style="width: 20%" align="right">
                </td>
                <td colspan="3">
                <asp:Label ID="lblEditCreditProduct0" runat="server" 
                        Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="PRODUCT"></asp:Label>&nbsp;<asp:DropDownList 
                        ID="ddlNewCreditProduct" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small"></asp:DropDownList>&nbsp;<asp:Label ID="Label32" 
                        runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" 
                        Text="Credit:"></asp:Label>&nbsp;<telerik:RadNumericTextBox 
                        ID="rntbNewCredit" runat="server" DataType="System.Int32" Font-Names="Verdana" 
                        Font-Size="XX-Small" MaxValue="999" MinValue="0" ShowSpinButtons="True" 
                        Width="60px"><NumberFormat DecimalDigits="0" ZeroPattern="n" /></telerik:RadNumericTextBox>&nbsp;<asp:CheckBox 
                        ID="cbNewCreditEnforce" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Enforce" TextAlign="Left" />&nbsp;<asp:Label 
                        ID="Label30" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="Start:"></asp:Label>&nbsp;<telerik:RadDateTimePicker 
                        ID="rdtpNewCreditStartDateTime" runat="server" Culture="en-GB" 
                        DateInput-DateFormat="dd-MMM-yyyy " DateInput-DisplayDateFormat="dd-MMM-yyyy " 
                        FocusedDate="2013-01-01" Font-Names="Verdana" Font-Size="XX-Small" 
                        MinDate="2013-01-01" Width="180px"><TimeView CellSpacing="-1" Culture="en-GB"></TimeView><TimePopupButton 
                            HoverImageUrl="" ImageUrl="" /><Calendar 
                            UseColumnHeadersAsSelectors="False" UseRowHeadersAsSelectors="False" 
                            ViewSelectorText="x"></Calendar><DateInput DateFormat="dd-MMM-yyyy hh:mm" 
                            DisplayDateFormat="dd-MMM-yyyy hh:mm" LabelWidth="40%"></DateInput><DatePopupButton 
                            HoverImageUrl="" ImageUrl="" /></telerik:RadDateTimePicker>&nbsp;<asp:Label 
                        ID="Label31" runat="server" Font-Bold="False" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="End:"></asp:Label>&nbsp;<telerik:RadDateTimePicker 
                        ID="rdtpNewCreditEndDateTime" runat="server" Culture="en-GB" 
                        DateInput-DateFormat="dd-MMM-yyyy " DateInput-DisplayDateFormat="dd-MMM-yyyy " 
                        FocusedDate="2013-01-01" Font-Names="Verdana" Font-Size="XX-Small" 
                        MinDate="2013-01-01" Width="180px"><TimeView CellSpacing="-1" Culture="en-GB"></TimeView><TimePopupButton 
                            HoverImageUrl="" ImageUrl="" /><Calendar 
                            UseColumnHeadersAsSelectors="False" UseRowHeadersAsSelectors="False" 
                            ViewSelectorText="x"></Calendar><DateInput DateFormat="dd-MMM-yyyy hh:mm" 
                            DisplayDateFormat="dd-MMM-yyyy hh:mm" LabelWidth="40%"></DateInput><DatePopupButton 
                            HoverImageUrl="" ImageUrl="" /></telerik:RadDateTimePicker>&nbsp;<asp:Button 
                        ID="btnNewCreditSave" runat="server" Text="save" 
                        onclick="btnNewCreditSave_Click" />&nbsp;<asp:Button 
                        ID="btnNewCreditCancel" runat="server" Text="cancel" 
                        onclick="btnNewCreditCancel_Click" /></td><td style="width: 1%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <!-- usage: onmouseover="return escape('Tooltip text')" -->
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
