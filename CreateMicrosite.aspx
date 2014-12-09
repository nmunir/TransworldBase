<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="SprintInternational" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.DirectoryServices" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    
    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8
    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16
    
    Const WEBSITE_NAME As String = "Default Web Site" '"my.transworld.eu.com"
    Const PHYSICAL_PATH As String = "c:\" '"d:\CourierSoftware\www\common"
    Const IMAGES_DIR As String = "prod_images"
    Const IMAGES_DIR_PHYSICAL_PATH As String = "c:\" '"d:\CourierSoftware\www\common\images"
    
    Protected Sub Page_Load()
        If Not IsPostBack Then
            Call RemoveLoginCookie()
            Call ShowSiteSetup()
            Call GetAccountHandlers()
        End If
    End Sub
    
    Protected Sub GetAccountHandlers()
        Dim items As ListItemCollection = ExecuteQueryToListItemCollection("SELECT * FROM AccountHandler WHERE DeletedFlag = 0", "name", "key")
        ddlAccountHandlers.Items.Clear()
        ddlAccountHandlers.Items.Add(New ListItem("- select an account handler -","-1"))
        For Each li As ListItem In items
            ddlAccountHandlers.Items.Add(li)
        Next
        ddlAccountHandlers.SelectedValue = -1
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
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
        'tbSiteName.Text = sGetPath()
    End Sub
    
    Protected Sub HideAllPanels()
        'pnlMain.Visible = False
        'pnlMustChangePassword.Visible = False
        'pnlSiteSetup.Visible = False
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
    
    Protected Sub btnContinueWithDefault_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("SiteKey") = 0
        Server.Transfer("session_expired.aspx")
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
    
    Protected Sub TrimFields()
        tbFirstName.Text = tbFirstName.Text.Trim
        tbLastName.Text = tbLastName.Text.Trim
        tbUserId.Text = tbUserId.Text.Trim
        tbPassword.Text = tbPassword.Text.Trim
        tbEmailAddr.Text = tbEmailAddr.Text.Trim
    End Sub
    
    Protected Sub btnSaveSiteMapping_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TrimFields()
        Call SaveSiteConfiguration()
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
    
    Protected Function sitenameAlreadyExistsForTheCustomer(ByVal customerKey As String) As Boolean
        Return ExecuteQueryToDataTable("select * from sitenametokeymap where sitekey = " & customerKey).Rows.Count > 0
    End Function
    
    Protected Function useridAlreadyTaken(ByVal userid As String) As Boolean
        Return ExecuteQueryToDataTable("select * from userprofile where userid = '" & userid & "'").Rows.Count > 0
    End Function
    
    Protected Function sitenameAlreadyTaken(ByVal siteName As String) As Boolean
        Return ExecuteQueryToDataTable("select * from sitenametokeymap where path = '" & siteName & "'").Rows.Count > 0
    End Function

    Protected Function sitenamesAlreadyTaken(ByVal siteNames As String) As Boolean
        sitenamesAlreadyTaken = True
        Dim listOfAliases As String() = siteNames.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim actual As String = listOfAliases(count).Trim()
            sitenamesAlreadyTaken = sitenamesAlreadyTaken and sitenameAlreadyTaken(actual)
        Next
    End Function
    
    Protected Function accountHandlerExistsForTheCustomer(ByVal customerkey As String) As Boolean
        Return ExecuteQueryToDataTable("select * from userprofile where userpermissions = 11 and customer = 0 and customerkey = " & customerkey).Rows.Count > 0
    End Function
    
    Protected Function formIsValid() As Boolean
        Dim bResult As Boolean = True
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim actual As String = listOfAliases(count).Trim()
            Dim replaced As String = actual.Replace("-", "")
            If actual.Length - replaced.Length > 1 Or actual.Length > 16 Or Not New Regex("^[A-Za-z0-9]+$").IsMatch(replaced) Then
                WebMsgBox.Show("The Site Name should : \n   Contain alphanumeric characters\n   Contain at most one hyphen\n   Not be longer than 16 characters.")
                bResult = False
                Exit For
            End If
        Next
        If ddlCustomerAccountCodes.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select a customer name.")
            bResult = False
        End If
        
        If sitenameAlreadyTaken(tbSiteName.Text) Then
            WebMsgBox.Show("The Site Name specified is already in use - cannot continue")
            bResult = False
        End If
        If ddlAccountHandlers.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select an account handler.")
            bResult = False
        End If
        If tbFirstName.Text = "" Then
            WebMsgBox.Show("Please enter a valid First Name.\n   Cannot be empty \n   Cannot be longer than 50 characters.")
            bResult = False
        End If
        If tbLastName.Text = "" Then
            WebMsgBox.Show("Please enter a Last Name.\n   Cannot be empty \n   Cannot be longer than 50 characters.")
            bResult = False
        End If
        If tbUserId.Text = "" Then
            WebMsgBox.Show("Please enter a User Id.\n   Cannot be empty \n   Cannot be longer than 100 characters.")
            bResult = False
        End If
        If tbPassword.Text = "" Then
            WebMsgBox.Show("Please enter a Password.\n   Cannot be empty \n   Cannot be longer than 24 characters.")
            bResult = False
        End If
        If Not New Regex("^([a-z0-9_\.-]+)@([\da-z\.-]+)\.([a-z\.]{2,6})$").IsMatch(tbEmailAddr.Text) Then
            WebMsgBox.Show("Please enter a valid Email Address.\n   Cannot be empty \n   Cannot be longer than 100 characters.")
            bResult = False
        End If
        If bUserIdExists(tbUserId.Text) Then
            WebMsgBox.Show("The UserId specified already exists - cannot continue")
            bResult = False
        End If
        If ddlCustomerAccountCodes.SelectedIndex < 2 Then
            WebMsgBox.Show("Please select a customer from the dropdown.")
            bResult = False
        End If
        If useridAlreadyTaken(tbUserId.Text) Then
            WebMsgBox.Show("The entered User Id is already taken. Please enter another.")
            bResult = False
        End If
        If sitenamesAlreadyTaken(tbSiteName.Text) Then
            WebMsgBox.Show("The entered Site Name(s) is already taken. Please enter another.")
            bResult = False
        End If
        Dim warningText As String = ""
        If accountHandlerExistsForTheCustomer(ddlCustomerAccountCodes.SelectedValue) Then
            warningText = warningText & "WARNING: Account Handler's account already exists for this customer. Continuing...\n"
        End If
        If sitenameAlreadyExistsForTheCustomer(ddlCustomerAccountCodes.SelectedValue) Then
            warningText = warningText & "WARNING: There already exists a Site Name entry for this customer. Continuing..."
        End If
        If warningText.Length > 0 Then
            WebMsgBox.Show(warningText)
        End If
        Return bResult
    End Function
    
    Protected Function getWebsiteDirectory(ByVal websiteName As String) As DirectoryEntry
        Dim root As New DirectoryEntry("IIS://localhost/W3SVC", "", "")
        Dim website As DirectoryEntry = Nothing
        For Each de As DirectoryEntry In root.Children   ' Access is denied.
            If de.SchemaClassName = "IIsWebServer" Then
                If de.Properties("ServerComment").Value.ToString() = websiteName Then
                    website = de
                End If
            End If
        Next
        If IsNothing(website) Then
            Return Nothing
        End If
        For Each d As DirectoryEntry In website.Children
            If d.Name = "ROOT" Then
                website = d
            End If
        Next
        Return website
    End Function
    
    Protected Function getWebsiteSubdirectory(ByVal parentDir As DirectoryEntry, ByVal subDirName As String) As DirectoryEntry
        For Each d As DirectoryEntry In parentDir.Children
            If d.Name = subDirName Then
                Return d
            End If
        Next
        Return Nothing
    End Function
    
    Protected Function addApplicationToWebsite(ByVal website As DirectoryEntry, ByVal appName As String, ByVal physicalPath As String) As Boolean
        Dim bRet As Boolean = True
        Dim newVdir As DirectoryEntry = Nothing
        
        Try
            newVdir = website.Children.Add(appName, "IIsWebVirtualDir")
            newVdir.CommitChanges()

            newVdir.Properties("Path").Item(0) = physicalPath
            newVdir.Properties("AccessRead").Item(0) = True
            newVdir.Properties("AccessWrite").Item(0) = True
            newVdir.Properties("EnableDirBrowsing").Item(0) = False
            newVdir.Properties("AccessScript").Item(0) = True
            newVdir.Properties("AppFriendlyName").Item(0) = appName
            newVdir.Properties("AppIsolated").Item(0) = 2
            newVdir.Invoke("AppCreate2", 0)

            newVdir.CommitChanges()
            website.CommitChanges()
        Catch e As Exception
            bRet = False
            WebMsgBox.Show("EXCEPTION... Virtual Directory Creation... Message : " & e.Message)
        Finally
            If Not IsNothing(newVdir) Then
                newVdir.Close()
            End If
            If Not IsNothing(website) Then
                website.Close()
            End If
        End Try
        
        Return bRet
    End Function
    
    Protected Function addVirtualDirectoryToDirectory(ByVal directory As DirectoryEntry, ByVal virDirName As String, ByVal physicalPath As String) As Boolean
        Dim bRet As Boolean = True
        Dim newVdir As DirectoryEntry = Nothing
        Try
            newVdir = directory.Children.Add(virDirName, "IIsWebVirtualDir")
            newVdir.CommitChanges()

            newVdir.Properties("Path").Item(0) = physicalPath
            newVdir.Properties("AccessRead").Item(0) = True
            newVdir.Properties("AccessWrite").Item(0) = True
            newVdir.Properties("EnableDirBrowsing").Item(0) = False
            newVdir.Properties("AccessScript").Item(0) = True

            newVdir.CommitChanges()
            directory.CommitChanges()
        Catch e As Exception
            bRet = False
            WebMsgBox.Show("EXCEPTION... Virtual Directory Creation... Message : " & e.Message)
        Finally
            If Not IsNothing(newVdir) Then
                newVdir.Close()
            End If
            If Not IsNothing(directory) Then
                directory.Close()
            End If
        End Try
        Return bRet
    End Function
    
    Protected Function CreateVirtualDir() As Boolean
        Dim website As DirectoryEntry = getWebsiteDirectory(WEBSITE_NAME)
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim virDirName As String = listOfAliases(count).Trim()
        
            If IsNothing(website) Then
                WebMsgBox.Show("No website found by that name. Please verify if it is correct.")
                Return False
            End If
        
            If Not IsNothing(getWebsiteSubdirectory(website, virDirName)) Then ' means that a directory by than name already exists ERROR!!!
                WebMsgBox.Show("Folder by the name '" & virDirName & "' already exists.")
                Return False
            End If
        
            If Not addApplicationToWebsite(website, virDirName, PHYSICAL_PATH) Then
                Return False
            End If
        
            If Not addVirtualDirectoryToDirectory(getWebsiteSubdirectory(website, virDirName), IMAGES_DIR, IMAGES_DIR_PHYSICAL_PATH) Then
                Return False
            End If
        Next
        Return True
    End Function
    
    Protected Sub SendEmailAlert(ByVal sRecipient As String, ByVal sSubject As String, ByVal sText As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = "MICROSITE_CREATION"
    
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

    Protected Function BuildEmailBody() As String
        BuildEmailBody = ""
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim actual As String = listOfAliases(count).Trim()
            BuildEmailBody = BuildEmailBody & "<br/>" & "A microsite by the alias '" & actual & "' has been created for the customer '" & ddlCustomerAccountCodes.SelectedItem.Text & "' and you have been assigned as the Account Handler."
        Next
    End Function

    Protected Function CreateSiteNameEntries() As Boolean
        CreateSiteNameEntries = True
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim actual As String = listOfAliases(count).Trim()
            CreateSiteNameEntries = CreateSiteNameEntries And ExecuteNonQuery("INSERT INTO SiteNameToKeyMap (Path, SiteKey) VALUES ('" & actual & "', " & CInt(ddlCustomerAccountCodes.SelectedValue) & ")")
        Next
    End Function

    Protected Sub RollbackEverything()
        RemoveVirtualDirectories()
        RemoveSiteNameEntries()
    End Sub

    Protected Sub RemoveVirtualDirectories()
        Dim website As DirectoryEntry = getWebsiteDirectory(WEBSITE_NAME)
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim virDirName As String = listOfAliases(count).Trim()
        
            If IsNothing(website) Then
                Continue For
            End If
        
            If Not IsNothing(getWebsiteSubdirectory(website, virDirName)) Then ' means that a directory by than name already exists ERROR!!!
                website.Children.Remove(getWebsiteSubdirectory(website, virDirName))
                website.CommitChanges()
            End If
        Next
    End Sub

    Protected Sub RemoveSiteNameEntries()
        Dim listOfAliases As String() = tbSiteName.Text.Split(",")
        For count As Integer = 0 To listOfAliases.Length - 1
            Dim actual As String = listOfAliases(count).Trim()
            ExecuteNonQuery("delete from SiteNameToKeyMap where path = '" & actual & "' and sitekey=" & CInt(ddlCustomerAccountCodes.SelectedValue))
        Next
    End Sub
        
    Protected Sub SaveSiteConfiguration()
        Dim success As Boolean = False
        If Not formIsValid() Then
            Exit Sub
        End If
        If CreateVirtualDir() Then
            If CreateSiteNameEntries() Then
                If AddNewUser() Then
                    success = True
                    SendEmailAlert(tbEmailAddr.Text, "Microsite Creation", BuildEmailBody())
                    WebMsgBox.Show("Site created successfully.")
                End If
            End If
        End If
        
        If Not success Then
            'rollbackEverything()
        End If

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
        
        'Session("SiteKey") = CInt(tbSiteKey.Text)
        
        'Dim sXMLRotatorConfigSourceFilePath As String
        'Dim sXMLNewsContentSourceFilePath As String
        'Dim sXMLRotatorConfigDestinationFilePath As String
        'Dim sXMLNewsContentDestinationFilePath As String
        'sXMLRotatorConfigSourceFilePath = ".\rotator\news_config0" & ".xml"
        'sXMLNewsContentSourceFilePath = ".\rotator\news0" & ".xml"
        'sXMLRotatorConfigDestinationFilePath = ".\rotator\news_config" & Session("SiteKey") & ".xml"
        'sXMLNewsContentDestinationFilePath = ".\rotator\news" & Session("SiteKey") & ".xml"
        'If Not My.Computer.FileSystem.FileExists(MapPath(sXMLRotatorConfigDestinationFilePath)) Then
        '    My.Computer.FileSystem.CopyFile(MapPath(sXMLRotatorConfigSourceFilePath), MapPath(sXMLRotatorConfigDestinationFilePath))
        'End If
        'If Not My.Computer.FileSystem.FileExists(MapPath(sXMLNewsContentDestinationFilePath)) Then
        '    My.Computer.FileSystem.CopyFile(MapPath(sXMLNewsContentSourceFilePath), MapPath(sXMLNewsContentDestinationFilePath))
        'End If
        'If cbAddDefaultUser.Checked Then
        '    Call AddNewUser()
        'End If
        'Server.Transfer("session_expired.aspx")
    End Sub
    
    Protected Function AddNewUser() As Boolean
        If Not accountHandlerExistsForTheCustomer(ddlCustomerAccountCodes.SelectedValue) Then
            Dim bError As Boolean
            tbUserId.Text = tbUserId.Text.Trim
            If tbUserId.Text.ToLower = "sa" Then
                WebMsgBox.Show("SA is a reserved User ID, please reselect")
                Return False
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
                    Return False
                Else
                    WebMsgBox.Show("AddNewUser: " & ex.ToString)
                    Return False
                End If
            Finally
                oConn.Close()
            End Try
        End If
        Return True
    End Function
    
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
    
    Protected Function bUserIdExists(ByVal sUserId As String) As Boolean
        bUserIdExists = False
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
    
    Protected Sub ddlAccountHandlers_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim selVal As Integer = ddlAccountHandlers.SelectedValue
        Dim data As DataTable = ExecuteQueryToDataTable("SELECT * FROM AccountHandler WHERE DeletedFlag = 0 and [key] = " & selVal)
        If data.Rows.Count > 0 Then
            Dim dr As DataRow = data.Rows(0)
            Dim fullName As String = safeDBGet("name", dr)
            Dim nameInParts As String() = fullName.Split(" ")
            tbLastName.Text = ""
            If nameInParts.Length > 0 Then
                tbFirstName.Text = nameInParts(0)
                tbUserId.Text = tbSiteName.Text & tbFirstName.Text
            End If
            For count As Integer = 1 To nameInParts.Length - 1
                tbLastName.Text = tbLastName.Text & nameInParts(count) & " "
            Next
            tbEmailAddr.Text = safeDBGet("emailaddr", dr)
        Else
            tbFirstName.Text = ""
            tbUserId.Text = ""
            tbLastName.Text = ""
            tbEmailAddr.Text = ""
            tbPassword.Text = ""
        End If
    End Sub
    
    Protected Function safeDBGet(ByVal col As String, ByVal dr As DataRow) As String
        If IsDBNull(dr(col)) Then
            Return ""
        End If
        Return dr(col).ToString().Trim()
    End Function
        
    Protected Sub SiteName_TextChanged()
        If ddlAccountHandlers.SelectedValue <> -1 Then
            tbUserId.Text = tbSiteName.Text & tbFirstName.Text
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
   <asp:Panel ID="pnlSiteSetup" Visible="false" runat="server" Width="100%">
            <asp:Label ID="Label1" runat="server" Text="Microsite Configuration for Customer Accounts" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label><br />
            <br />
            
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
                    
                        <asp:TextBox ID="tbSiteName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onkeypress="tbSiteName_keyPress(event)" onblur="document.getElementById('tbUserId').value=this.value+document.getElementById('tbFirstName').value;" AutoPostBack="true" CausesValidation="false"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvSiteName" runat="server" ControlToValidate="tbSiteName"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator>
                    </td>
                    
                </tr>
                <!--tr>
                    <td align="right">
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Site Key:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbSiteKey" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="tbSiteKey"
                            ErrorMessage="must be a number between 0 & 9999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="9999" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="rfvSiteKey" runat="server" ControlToValidate="tbSiteKey"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr-->
                <tr>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                </tr>
                <!--tr>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbAddDefaultUser" runat="server" Checked="True" Font-Names="Verdana"
                            Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbAddDefaultUser_CheckedChanged" Text="Add default user" />
                        &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
                    </td>
                </tr-->
                <tr>
                    <td align="right">
                    </td>
                    <td>
                        </td>
                </tr>

                <tr>
                    <td align="right">
                        <asp:Label ID="accountHandlersLabel" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Account Handlers" ></asp:Label>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlAccountHandlers" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" AutoPostBack="true" 
                            onselectedindexchanged="ddlAccountHandlers_SelectedIndexChanged">
                        
                        </asp:DropDownList>
                    </td>
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
                            Width="250px"></asp:TextBox>
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
                <!--tr>
                    <td align="right">
                        <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Router access:"></asp:Label></td>
                    <td>
                        <asp:TextBox ID="tbSecurityPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvSecurityPassword" runat="server" ControlToValidate="tbSecurityPassword"
                            ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator></td>
                </tr-->
                
                <tr>
                    <td></td>
                    <td align="left">
                        <asp:Button ID="btnSaveSiteMapping" runat="server" OnClick="btnSaveSiteMapping_Click"
                            Text="Create Site" /></td>
                    
                </tr>
            </table>
        </asp:Panel>
        </form>
</body>
</html>