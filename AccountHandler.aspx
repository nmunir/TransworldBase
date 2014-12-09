<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" MaintainScrollPositionOnPostback="true" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>

<script runat="server">

    ' TO DO

    ' NOT CURRENTLY DEALING WITH tbCollectionAddress2  !!!!!!!!!!!!!!!
    
    ' put in GKG changes for PWC
    ' when logging out from autologin, old username appears, with comma
    ' event.aspx: hide save changes button if < 2 days until start of event
    ' do user manual
    ' put in autodelete of old events (3 months to begin with)
    ' Custom report for calendar managed products
    ' Customer Visible?
    ' later on, allow existing event details to be copied to new event
    ' add help to multiple bookings
    ' no headers in on_line_picks.aspx list of items
    
    Const COUNTRY_KEY_UK As Integer = 222
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Const USER_PERMISSION_ADMINISTRATOR As Integer = 2

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            tbEmailReminderText.Attributes.Add("onkeypress", "return clickButton(event,'" + btnAddReminder.ClientID + "')")
            tbNote.Attributes.Add("onkeypress", "return clickButton(event,'" + btnAddNote.ClientID + "')")
            Call PopulateSiteAdministratorDropdown()
        Call GetSiteFeatures()
        
        End If
        Call SetTitle()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Account Handler"
    End Sub
    
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent3", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "SiteSettingsAndWebForm"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        cbPostcodeLookup.Checked = dr("PostcodeLookup")
        cbCalendarManagement.Checked = dr("CalendarManagement")
        cbUserPermissions.Checked = dr("UserPermissions")
        cbFileUpload.Checked = dr("FileUpload")
        If cbFileUpload.Checked Then
            lblLegendFileUploadNotificationEmailAddresses.Visible = True
            tbFileUploadNotificationEmailAddresses.Visible = True
            tbFileUploadNotificationEmailAddresses.Text = GetFileUploadNotificationEmailAddresses()
        Else
            lblLegendFileUploadNotificationEmailAddresses.Visible = False
            tbFileUploadNotificationEmailAddresses.Visible = False
        End If
        cbProductOwners.Checked = dr("ProductOwners")
        tbCategoryCount.Text = dr("CategoryCount")

        cbUseLabelPrinter.Checked = dr("UseLabelPrinter")
        cbSearchCompanyNameOnly.Checked = dr("SearchCompanyNameOnly")
        tbDefaultDescription.Text = dr("DefaultDescription")
        cbMakeRef1Mandatory.Checked = dr("MakeRef1Mandatory")
        tbRef1Label.Text = dr("Ref1Label")
        cbMakeRef2Mandatory.Checked = dr("MakeRef2Mandatory")
        tbRef2Label.Text = dr("Ref2Label")
        cbMakeRef3Mandatory.Checked = dr("MakeRef3Mandatory")
        tbRef3Label.Text = dr("Ref3Label")
        cbMakeRef4Mandatory.Checked = dr("MakeRef4Mandatory")
        tbRef4Label.Text = dr("Ref4Label")
        tbThirdPartyCollectionKey.Text = dr("ThirdPartyCollectionKey")
        cbHideCollectionButton.Checked = dr("HideCollectionButton")
        'cbPrintOndemandTab.Checked = dr("Misc1")
        cbOnDemandProducts.Checked = dr("OnDemandProducts")
        cbCustomLetters.Checked = dr("CustomLetters")
        Call GetAttemAccessParameters()
        'trAttemAccess.Visible = cbOnDemandProducts.Checked
        cbWebForm.Checked = dr("WebForm")
        If cbWebForm.Checked Then
            SetWebFormControlsVisibility(True)
        Else
            SetWebFormControlsVisibility(False)
        End If

        cbCustRef1IsVisible.Checked = dr("StockOrderCustRef1Visible")
        cbCustRef1IsMandatory.Checked = dr("StockOrderCustRef1Mandatory")
        tbCustRef1Label.Text = dr("StockOrderCustRefLabel1Legend")
        Call SetCustRef1Enabled(cbCustRef1IsVisible.Checked)

        cbCustRef2IsVisible.Checked = dr("StockOrderCustRef2Visible")
        cbCustRef2IsMandatory.Checked = dr("StockOrderCustRef2Mandatory")
        tbCustRef2Label.Text = dr("StockOrderCustRefLabel2Legend")
        Call SetCustRef2Enabled(cbCustRef2IsVisible.Checked)

        cbCustRef3IsVisible.Checked = dr("StockOrderCustRef3Visible")
        cbCustRef3IsMandatory.Checked = dr("StockOrderCustRef3Mandatory")
        tbCustRef3Label.Text = dr("StockOrderCustRefLabel3Legend")
        Call SetCustRef3Enabled(cbCustRef3IsVisible.Checked)

        cbCustRef4IsVisible.Checked = dr("StockOrderCustRef4Visible")
        cbCustRef4IsMandatory.Checked = dr("StockOrderCustRef4Mandatory")
        tbCustRef4Label.Text = dr("StockOrderCustRefLabel4Legend")
        Call SetCustRef4Enabled(cbCustRef4IsVisible.Checked)
        
        tbSessionTimeout.Text = dr("SessionTimeout")
        
        'tbWebFormCustomerKey.Text = dr("WebFormCustomerKey")
        If IsNumeric(dr("WebFormCustomerKey")) Then
            Call SetWebFormCustomer(CInt(dr("WebFormCustomerKey")))
        Else
            Call SetWebFormCustomer(0)
        End If
        
        'tbWebFormGenericUserKey.Text = dr("WebFormGenericUserKey")
        If IsNumeric(dr("WebFormCustomerKey")) Then
            If IsNumeric(dr("WebFormGenericUserKey")) Then
                Call SetWebFormGenericUser(CInt(dr("WebFormCustomerKey")), CInt(dr("WebFormGenericUserKey")))
            Else
                Call SetWebFormGenericUser(CInt(dr("WebFormCustomerKey")), 0)
            End If
        Else
            ddlWebFormGenericUser.Enabled = False
        End If

        tbWebFormPageTitle.Text = dr("WebFormPageTitle")
        tbWebFormLogoImage.Text = dr("WebFormLogoImage")
        tbWebFormTopLegend.Text = dr("WebFormTopLegend")
        tbWebFormBottomLegend.Text = dr("WebFormBottomLegend")
        cbWebFormShowPrice.Checked = dr("WebFormShowPrice")
        cbWebFormShowZeroQuantity.Checked = dr("WebFormShowZeroQuantity")
        cbWebFormZeroStockNotification.Checked = dr("WebFormZeroStockNotification")

        FCKedWebFormHomePage.Value = dr("WebFormHomePageText")
        FCKedWebFormAddressPage.Value = dr("WebFormAddressPageText")
        FCKedWebFormHelpPage.Value = dr("WebFormHelpPageText")
        
    End Sub
    
    Protected Function GetFileUploadNotificationEmailAddresses() As String
        GetFileUploadNotificationEmailAddresses = String.Empty
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT * FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey"), "EmailAddr", "id")
        For Each li As ListItem In oListItemCollection
            If Not GetFileUploadNotificationEmailAddresses = String.Empty Then
                GetFileUploadNotificationEmailAddresses += ", "
            End If
            
            GetFileUploadNotificationEmailAddresses += li.Text
        Next
    End Function
    
    Protected Sub SetWebFormCustomer(ByVal nWebFormCustomerKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND DeletedFlag = 'N' ORDER BY CustomerAccountCode"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        li = New ListItem("- please select -", 0)
        ddlWebFormCustomer.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("CustomerAccountCode"), oDataReader("CustomerKey"))
                ddlWebFormCustomer.Items.Add(li)
            End While
        Catch ex As Exception
            WebMsgBox.Show("SetWebFormCustomer: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        For i As Integer = 0 To ddlWebFormCustomer.Items.Count - 1
            If ddlWebFormCustomer.Items(i).Value = nWebFormCustomerKey Then
                ddlWebFormCustomer.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub SetWebFormGenericUser(ByVal nWebFormCustomerKey As Integer, ByVal nWebFormGenericUserKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT FirstName, LastName, UserId, [Key] FROM UserProfile WHERE CustomerKey = " & nWebFormCustomerKey & " AND Status = 'Active' AND DeletedFlag = 0 ORDER BY LastName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        ddlWebFormGenericUser.Items.Clear()
        li = New ListItem("- please select -", 0)
        ddlWebFormGenericUser.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("FirstName") & " " & oDataReader("LastName") & " (" & oDataReader("UserId") & ")", oDataReader("Key"))
                ddlWebFormGenericUser.Items.Add(li)
            End While
        Catch ex As Exception
            WebMsgBox.Show("SetWebFormGenericUser: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        For i As Integer = 0 To ddlWebFormGenericUser.Items.Count - 1
            If ddlWebFormGenericUser.Items(i).Value = nWebFormGenericUserKey Then
                ddlWebFormGenericUser.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub GetAttemAccessParameters()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.CommandText = "spASPNET_OnDemand_AccessGetParameters"

        oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")

        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                tbAttemUserName.Text = oDataReader("UserName")
                tbAttemPassword.Text = oDataReader("Password")
                tbAttemCustomerName.Text = oDataReader("Customer")
            Else
                tbAttemUserName.Text = String.Empty
                tbAttemPassword.Text = String.Empty
                tbAttemCustomerName.Text = String.Empty
            End If
            lblCustomerName.Text = Session("CustomerName")
        Catch ex As Exception
            WebMsgBox.Show("Error in GetAttemAccessParameters: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub PopulateSiteAdministratorDropdown()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT FirstName + ' ' + LastName + ' (' + UserId + ')' Name, [key], UserPermissions FROM UserProfile WHERE Type LIKE 'SuperUser' AND Status LIKE 'active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY FirstName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Dim li As ListItem
        li = New ListItem("- please select -", 0)
        ddlSiteAdministrator.Items.Add(li)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                li = New ListItem(oDataReader("Name"), oDataReader("Key"))
                ddlSiteAdministrator.Items.Add(li)
                If Not IsDBNull(oDataReader("UserPermissions")) Then
                    If oDataReader("UserPermissions") And USER_PERMISSION_ADMINISTRATOR Then
                        ddlSiteAdministrator.SelectedIndex = ddlSiteAdministrator.Items.Count - 1
                        hidSiteAdministratorKey.Value = oDataReader("Key")
                    End If
                End If
            End While
        Catch ex As Exception
            WebMsgBox.Show("PopulateSiteAdministratorDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnShowEvents_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Call HideAllPanels()
        Call ShowEvents(bEventTypeAll:=b.CommandArgument.ToLower = "all", bCustomerTypeAll:=rbCMAllCustomers.Checked)
        lblItemsPerPage.Visible = True
        ddlCMItemsPerPage.Visible = True
    End Sub

    Protected Sub ShowEvents(ByVal bEventTypeAll As Boolean, ByVal bCustomerTypeAll As Boolean)   ' false=events awaiting review / true=all events, false=my customers / true = all customers
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        If bEventTypeAll Then
            If bCustomerTypeAll Then
                oCmd.CommandText = "spASPNET_CalendarManaged_GetEventsAllForAccountHandlerAllCustomers"
                pnDisplayMode = 1
                lblLegendEvents.Text = "All Events (all customers)"
            Else
                oCmd.CommandText = "spASPNET_CalendarManaged_GetEventsAllForAccountHandlerAHCustomers"
                pnDisplayMode = 2
                lblLegendEvents.Text = "All Events (" & ddlAccountHandler.SelectedItem.Text & "'s customers)"
            End If
        Else
            If bCustomerTypeAll Then
                oCmd.CommandText = "spASPNET_CalendarManaged_GetEventsUnreviewedForAccountHandlerAllCustomers"
                pnDisplayMode = 3
                lblLegendEvents.Text = " Events Awaiting Review (all customers)"
            Else
                oCmd.CommandText = "spASPNET_CalendarManaged_GetEventsUnreviewedForAccountHandlerAHCustomers"
                pnDisplayMode = 4
                lblLegendEvents.Text = " Events Awaiting Review (" & ddlAccountHandler.SelectedItem.Text & "'s customers)"
            End If
        End If
        oCmd.CommandText += "4"
        psRefreshStoredProcedure = oCmd.CommandText
        If Not bCustomerTypeAll Then
            Dim paramAccountHandlerKey As SqlParameter = New SqlParameter("@AccountHandlerKey", SqlDbType.Int)
            paramAccountHandlerKey.Value = ddlAccountHandler.SelectedValue
            oCmd.Parameters.Add(paramAccountHandlerKey)
        End If
        
        Dim paramRetrospectiveDays As SqlParameter = New SqlParameter("@RetrospectiveDays", SqlDbType.Int)
        If cbRetrospective.Checked Then
            paramRetrospectiveDays.Value = ddlCMRetrospectiveDays.SelectedValue
        Else
            paramRetrospectiveDays.Value = 1
        End If
        oCmd.Parameters.Add(paramRetrospectiveDays)
        
        Dim paramIncludeDeleted As SqlParameter = New SqlParameter("@DeletedEvents", SqlDbType.Bit)
        If cbIncludeCancelledEvents.Checked Then
            paramIncludeDeleted.Value = 1
        Else
            paramIncludeDeleted.Value = 0
        End If
        oCmd.Parameters.Add(paramIncludeDeleted)
        
        oCmd.CommandText = oCmd.CommandText
        Dim bEventType = bEventTypeAll
        Dim bCustomerType = bCustomerTypeAll
        
        gvEvents.PageSize = ddlCMItemsPerPage.SelectedValue
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            Dim arr As ArrayList = New ArrayList
            For Each row As Object In oDataReader
                arr.Add(row)
            Next
            gvEvents.DataSource = arr
            gvEvents.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in ShowEvents: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        gvEvents.Visible = True
        lblLegendEvents.Visible = True
    End Sub
    
    Protected Sub RefreshEvents()
        If psRefreshStoredProcedure.ToLower.Contains("not defined") Then
            Exit Sub
        End If
        
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.CommandText = psRefreshStoredProcedure
        If psRefreshStoredProcedure.ToLower.Contains("ahcustomer") Then
            Dim paramAccountHandlerKey As SqlParameter = New SqlParameter("@AccountHandlerKey", SqlDbType.Int)
            paramAccountHandlerKey.Value = ddlAccountHandler.SelectedValue
            oCmd.Parameters.Add(paramAccountHandlerKey)
        End If

        Dim paramRetrospectiveDays As SqlParameter = New SqlParameter("@RetrospectiveDays", SqlDbType.Int)
        If cbRetrospective.Checked Then
            paramRetrospectiveDays.Value = ddlCMRetrospectiveDays.SelectedValue
        Else
            paramRetrospectiveDays.Value = 0
        End If
        oCmd.Parameters.Add(paramRetrospectiveDays)
        
        Dim paramIncludeDeleted As SqlParameter = New SqlParameter("@DeletedEvents", SqlDbType.Bit)
        If cbIncludeCancelledEvents.Checked Then
            paramIncludeDeleted.Value = 1
        Else
            paramIncludeDeleted.Value = 0
        End If
        oCmd.Parameters.Add(paramIncludeDeleted)
        
        gvEvents.PageSize = ddlCMItemsPerPage.SelectedValue

        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            Dim arr As ArrayList = New ArrayList
            For Each row As Object In oDataReader
                arr.Add(row)
            Next
            gvEvents.DataSource = arr
            gvEvents.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in RefreshEvents: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub InitAccountHandlerDropdown()
        If ddlAccountHandler.Items.Count = 0 Then
            Call PopulateAccountHandlerDropdown()
        End If
        ' do selection from cookie here
    End Sub
    
    Protected Sub PopulateAccountHandlerDropdown()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spStockMngr_AccountHandler_GetAll", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        Dim arr As ArrayList = New ArrayList
        ddlAccountHandler.Items.Add(New ListItem("- please select -", -1))
        For Each row As Object In oDataReader
            ddlAccountHandler.Items.Add(New ListItem(oDataReader("Name"), oDataReader("Key")))
        Next
        oConn.Close()
    End Sub

    Protected Sub ddlAccountHandler_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            btnShowAllEvents.Enabled = True
            btnShowEventsAwaitingReview.Enabled = True
            lblLegendEvents.Visible = False
            gvEvents.Visible = False
        End If
        ' do store of selection to cookie here, or clear cookie if index(0) selected
    End Sub

    Protected Sub rbCMAllCustomers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            lblLegendSelectAccountHandler.Visible = False
            ddlAccountHandler.Visible = False
            lblLegendEvents.Visible = False
            gvEvents.Visible = False
            btnShowAllEvents.Enabled = True
            btnShowEventsAwaitingReview.Enabled = True
        End If
    End Sub

    Protected Sub rbCMMyCustomers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            lblLegendSelectAccountHandler.Visible = False
            ddlAccountHandler.Visible = False
            lblLegendEvents.Visible = False
            gvEvents.Visible = False
            btnShowAllEvents.Enabled = False
            btnShowEventsAwaitingReview.Enabled = False
            lblLegendSelectAccountHandler.Visible = True
            ddlAccountHandler.Visible = True
            Call InitAccountHandlerDropdown()
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlEvent.Visible = False
    End Sub
    
    Protected Sub ShowEventPanel()
        Call HideAllPanels()
        pnlEvent.Visible = True
    End Sub
    
    Protected Sub lnkbtnReviewEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        pnEventId = CInt(lb.CommandArgument)
        Call InitCountryDropdowns()
        Call GetEventFromId()
        Call GetNotes()
        Call GetReminders()
        Call GetAccountHandlerEmailAddr()
        Call ShowEventPanel()
    End Sub

    Protected Sub GetNotes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventNotes2", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerVisibleOnly", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@CustomerVisibleOnly").Value = 0
        
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            gvNotes.DataSource = oDataTable
            gvNotes.DataBind()
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetReminders()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventReminders2", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            gvReminders.DataSource = oDataTable
            gvReminders.DataBind()
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetAccountHandlerEmailAddr()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_CalendarManaged_GetAccountHandlerEmailAddr", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                tbEmailReminderAddr.Text = dr("EmailAddr") & ","
            End If
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetEventFromId()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable1 As New DataTable
        Dim oAdapter1 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventById", oConn)
        Dim nDDLIndex As Integer
        Dim dtProductPickDate As DateTime = DateTime.MinValue

        oAdapter1.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter1.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter1.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        Try
            oConn.Open()
            oAdapter1.Fill(oDataTable1)
            Dim dr As DataRow = oDataTable1.Rows(0)
            ' need to check 1 and only 1 row present
            lblCustomer.Text = dr("CustomerAccountCode")
            hidCustomerKey.Value = dr("CustomerKey")
            lblEventName.Text = dr("EventName")
            tbContactName.Text = dr("ContactName")
            tbContactPhone.Text = dr("ContactPhone")
            tbContactMobile.Text = dr("ContactMobile")

            Dim sContactName2 As String
            Dim sContactPhone2 As String
            Dim sContactMobile2 As String
            
            If Not IsDBNull(dr("ProductPickDate")) Then
                dtProductPickDate = DateTime.Parse(dr("ProductPickDate"))
                lblProductPickedFlag.Text = "PICKED " & Format(dtProductPickDate, "d-MMM-yyyy h:mm")
            Else
                lblProductPickedFlag.Text = String.Empty
            End If

            If Not IsDBNull(dr("ContactName2")) Then
                sContactName2 = dr("ContactName2").ToString.Trim
            Else
                sContactName2 = String.Empty
            End If
            If Not IsDBNull(dr("ContactPhone2")) Then
                sContactPhone2 = dr("ContactPhone2").ToString.Trim
            Else
                sContactPhone2 = String.Empty
            End If
            If Not IsDBNull(dr("ContactMobile2")) Then
                sContactMobile2 = dr("ContactMobile2").ToString.Trim
            Else
                sContactMobile2 = String.Empty
            End If
            
            tbCMContactName2.Text = sContactName2
            tbCMContactPhone2.Text = sContactPhone2
            tbCMContactMobile2.Text = sContactMobile2
            If String.IsNullOrEmpty(sContactName2) And String.IsNullOrEmpty(sContactPhone2) And String.IsNullOrEmpty(sContactMobile2) Then
                Call SetContact2FieldsVisibility(False)
            Else
                Call SetContact2FieldsVisibility(True)
            End If

            tbEventAddress1.Text = dr("EventAddress1")
            tbEventAddress2.Text = dr("EventAddress2")
            tbEventAddress3.Text = dr("EventAddress3")
            tbTown.Text = dr("Town")
            tbPostcode.Text = dr("Postcode")

            Dim nCountryKey As Integer
            If Not IsDBNull(dr("CountryKey")) Then
                nCountryKey = dr("CountryKey")
            Else
                nCountryKey = COUNTRY_KEY_UK
            End If

            If nCountryKey = COUNTRY_KEY_UK Then
                trCMCountry.Visible = False
                lnkbtnCMAddressOutsideUK.Visible = True
            Else
                trCMCountry.Visible = True
                lnkbtnCMAddressOutsideUK.Visible = False
            End If

            For nDDLIndex = 1 To ddlCMCountry.Items.Count - 1
                If ddlCMCountry.Items(nDDLIndex).Value = nCountryKey Then
                    ddlCMCountry.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            
            Dim sTemp As String = dr("DeliveryTime")
            For nDDLIndex = 0 To ddlDeliveryTime.Items.Count - 1
                If ddlDeliveryTime.Items(nDDLIndex).Text = sTemp Then
                    ddlDeliveryTime.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            tbPreciseDeliveryPoint.Text = dr("PreciseDeliveryPoint")
            tbPreciseCollectionPoint.Text = dr("PreciseCollectionPoint")
            sTemp = dr("CollectionTime")
            For nDDLIndex = 0 To ddlCollectionTime.Items.Count - 1
                If ddlCollectionTime.Items(nDDLIndex).Text = sTemp Then
                    ddlCollectionTime.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            If Not IsDBNull(dr("DifferentCollectionAddress")) Then
                cbDifferentCollectionAddress.Checked = dr("DifferentCollectionAddress")
            Else
                cbDifferentCollectionAddress.Checked = False
            End If
            SetCollectionAddressVisibility(cbDifferentCollectionAddress.Checked)
            If cbDifferentCollectionAddress.Checked Then
                tbCollectionAddress1.Text = dr("CollectionAddress1")
                tbCollectionAddress2.Text = dr("CollectionAddress2")
                tbCollectionTown.Text = dr("CollectionTown")
                tbCollectionPostcode.Text = dr("CollectionPostcode")

                Dim nCollectionCountryKey As Integer
                If Not IsDBNull(dr("CountryKey")) Then
                    nCollectionCountryKey = dr("CountryKey")
                Else
                    nCollectionCountryKey = COUNTRY_KEY_UK
                End If

                If nCollectionCountryKey = COUNTRY_KEY_UK Then
                    trCMCollectionCountry.Visible = False
                    lnkbtnCMCollectionAddressOutsideUK.Visible = True
                Else
                    trCMCollectionCountry.Visible = True
                    lnkbtnCMCollectionAddressOutsideUK.Visible = False
                End If

                For nDDLIndex = 1 To ddlCMCollectionCountry.Items.Count - 1
                    If ddlCMCollectionCountry.Items(nDDLIndex).Value = nCollectionCountryKey Then
                        ddlCMCollectionCountry.SelectedIndex = nDDLIndex
                        Exit For
                    End If
                Next
            
            End If
            If Not IsDBNull(dr("CustomerReference")) Then
                tbCustomerReference.Text = dr("CustomerReference")
            Else
                tbCustomerReference.Text = String.Empty
            End If
            tbSpecialInstructions.Text = dr("SpecialInstructions")
            lblBookedBy.Text = dr("username")
            lblBookedOn.Text = dr("BookedOn")
            
            Dim oAdapter2 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventItemsById", oConn)
            
            oAdapter2.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter2.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
            oAdapter2.SelectCommand.Parameters("@EventId").Value = pnEventId
            Dim oDataTable2 As New DataTable
            oAdapter2.Fill(oDataTable2)
            gvItems.DataSource = oDataTable2
            gvItems.DataBind()
            
            If gvItems.Rows.Count = 1 Then
                lblLegendProduct.Text = "Product:"
                btnPickEventProducts.Text = "pick product for event"
            Else
                lblLegendProduct.Text = "Products:"
                btnPickEventProducts.Text = "pick products for event"
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try

        For Each gvr As GridViewRow In gvEvents.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim lb As LinkButton = gvr.FindControl("lnkbtnReviewEvent")
                If lb.CommandArgument = pnEventId Then
                    lblDeliveryDate.Text = gvr.Cells(7).Text
                    lblCollectionDate.Text = gvr.Cells(8).Text
                End If
            End If
        Next
    End Sub
    
    Protected Sub lnkbtnDeleteEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call DeleteEvent(CInt(lb.CommandArgument))
        pnlEvent.Visible = False
    End Sub

    Protected Sub DeleteEvent(ByVal nEventId As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_DeleteEvent", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = nEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.Int)
        paramUserId.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserId)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
        Call RefreshEvents()
    End Sub
    
    Protected Sub gvEvents_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.Cells.Count >= 3 Then      ' check if one or more rows - if no rows there will only be a single cell with the empty grid message
            row.Cells(1).Visible = False  ' hide items required in the query but not to be displayed (EventId, BookedBy)
            row.Cells(2).Visible = False
            row.Cells(3).Visible = False
            If IsNumeric(row.Cells(2).Text) Then
                If CInt(row.Cells(2).Text) = 1 Then
                    row.BackColor = Drawing.Color.LightGreen
                End If
            End If
        End If
    End Sub
    
    Protected Sub btnAddNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNote.Text = tbNote.Text.Trim
        If tbNote.Text.Length = 0 Then
            WebMsgBox.Show("Cannot add empty note!")
        Else
            Call AddNote(tbNote.Text, cbCustomerVisible.Checked)
            Call GetNotes()
        End If
        tbNote.Text = String.Empty
        cbCustomerVisible.Checked = True
    End Sub
    
    Protected Sub AddNote(ByVal sNoteText As String, ByVal bCustomerVisible As Boolean)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_AddNote", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.Int)
        paramUserId.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserId)

        Dim paramNote As SqlParameter = New SqlParameter("@Note", SqlDbType.NVarChar, 200)
        paramNote.Value = sNoteText
        oCmd.Parameters.Add(paramNote)

        Dim paramCustomerVisible As SqlParameter = New SqlParameter("@CustomerVisible", SqlDbType.Int)
        If bCustomerVisible Then
            paramCustomerVisible.Value = 1
        Else
            paramCustomerVisible.Value = 0
        End If
        oCmd.Parameters.Add(paramCustomerVisible)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
        tbNote.Text = String.Empty
    End Sub
    
    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("CalendarManaged")
        If Page.IsValid Then
            Call SaveEventChanges()
            WebMsgBox.Show("Your changes were saved.")
        Else
            WebMsgBox.Show("One or more fields were incorrect or not supplied. Please correct the information and resubmit.")
        End If
    End Sub
    
    Protected Sub SaveEventChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_UpdateEvent3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.VarChar, 50)
        paramContactName.Value = tbContactName.Text
        oCmd.Parameters.Add(paramContactName)

        Dim paramContactPhone As SqlParameter = New SqlParameter("@ContactPhone", SqlDbType.VarChar, 50)
        paramContactPhone.Value = tbContactPhone.Text
        oCmd.Parameters.Add(paramContactPhone)

        Dim paramContactMobile As SqlParameter = New SqlParameter("@ContactMobile", SqlDbType.VarChar, 50)
        paramContactMobile.Value = tbContactMobile.Text
        oCmd.Parameters.Add(paramContactMobile)

        Dim paramContactName2 As SqlParameter = New SqlParameter("@ContactName2", SqlDbType.VarChar, 50)
        paramContactName2.Value = tbCMContactName2.Text
        oCmd.Parameters.Add(paramContactName2)

        Dim paramContactPhone2 As SqlParameter = New SqlParameter("@ContactPhone2", SqlDbType.VarChar, 50)
        paramContactPhone2.Value = tbCMContactPhone2.Text
        oCmd.Parameters.Add(paramContactPhone2)

        Dim paramContactMobile2 As SqlParameter = New SqlParameter("@ContactMobile2", SqlDbType.VarChar, 50)
        paramContactMobile2.Value = tbCMContactMobile2.Text
        oCmd.Parameters.Add(paramContactMobile2)

        Dim paramEventAddress1 As SqlParameter = New SqlParameter("@EventAddress1", SqlDbType.VarChar, 50)
        paramEventAddress1.Value = tbEventAddress1.Text
        oCmd.Parameters.Add(paramEventAddress1)

        Dim paramEventAddress2 As SqlParameter = New SqlParameter("@EventAddress2", SqlDbType.VarChar, 50)
        paramEventAddress2.Value = tbEventAddress2.Text
        oCmd.Parameters.Add(paramEventAddress2)

        Dim paramEventAddress3 As SqlParameter = New SqlParameter("@EventAddress3", SqlDbType.VarChar, 50)
        paramEventAddress3.Value = tbEventAddress3.Text
        oCmd.Parameters.Add(paramEventAddress3)

        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.VarChar, 50)
        paramTown.Value = tbTown.Text
        oCmd.Parameters.Add(paramTown)

        Dim paramPostcode As SqlParameter = New SqlParameter("@Postcode", SqlDbType.VarChar, 50)
        paramPostcode.Value = tbPostcode.Text
        oCmd.Parameters.Add(paramPostcode)

        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int)
        paramCountryKey.Value = ddlCMCountry.SelectedValue
        oCmd.Parameters.Add(paramCountryKey)

        Dim paramDeliveryTime As SqlParameter = New SqlParameter("@DeliveryTime", SqlDbType.VarChar, 50)
        paramDeliveryTime.Value = ddlDeliveryTime.SelectedItem.Text
        oCmd.Parameters.Add(paramDeliveryTime)

        Dim paramPreciseDeliveryPoint As SqlParameter = New SqlParameter("@PreciseDeliveryPoint", SqlDbType.VarChar, 100)
        paramPreciseDeliveryPoint.Value = tbPreciseDeliveryPoint.Text
        oCmd.Parameters.Add(paramPreciseDeliveryPoint)

        Dim paramDifferentCollectionAddress As SqlParameter = New SqlParameter("@DifferentCollectionAddress", SqlDbType.Bit)
        paramDifferentCollectionAddress.Value = cbDifferentCollectionAddress.Checked
        oCmd.Parameters.Add(paramDifferentCollectionAddress)

        Dim paramCollectionAddress1 As SqlParameter = New SqlParameter("@CollectionAddress1", SqlDbType.NVarChar, 50)
        paramCollectionAddress1.Value = tbCollectionAddress1.Text
        oCmd.Parameters.Add(paramCollectionAddress1)

        Dim paramCollectionAddress2 As SqlParameter = New SqlParameter("@CollectionAddress2", SqlDbType.NVarChar, 50)
        paramCollectionAddress2.Value = tbCollectionAddress2.Text
        oCmd.Parameters.Add(paramCollectionAddress2)

        Dim paramCollectionTown As SqlParameter = New SqlParameter("@CollectionTown", SqlDbType.NVarChar, 50)
        paramCollectionTown.Value = tbCollectionTown.Text
        oCmd.Parameters.Add(paramCollectionTown)

        Dim paramCollectionPostcode As SqlParameter = New SqlParameter("@CollectionPostcode", SqlDbType.NVarChar, 50)
        paramCollectionPostcode.Value = tbCollectionPostcode.Text
        oCmd.Parameters.Add(paramCollectionPostcode)

        Dim paramCollectionCountryKey As SqlParameter = New SqlParameter("@CollectionCountryKey", SqlDbType.Int)
        paramCollectionCountryKey.Value = ddlCMCollectionCountry.SelectedValue
        oCmd.Parameters.Add(paramCollectionCountryKey)

        Dim paramCollectionTime As SqlParameter = New SqlParameter("@CollectionTime", SqlDbType.VarChar, 50)
        paramCollectionTime.Value = ddlCollectionTime.SelectedItem.Text
        oCmd.Parameters.Add(paramCollectionTime)

        Dim paramPreciseCollectionPoint As SqlParameter = New SqlParameter("@PreciseCollectionPoint", SqlDbType.VarChar, 100)
        paramPreciseCollectionPoint.Value = tbPreciseCollectionPoint.Text
        oCmd.Parameters.Add(paramPreciseCollectionPoint)

        Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.VarChar, 200)
        paramSpecialInstructions.Value = tbSpecialInstructions.Text
        oCmd.Parameters.Add(paramSpecialInstructions)

        Dim paramCustomerReference As SqlParameter = New SqlParameter("@CustomerReference", SqlDbType.NVarChar, 100)
        paramCustomerReference.Value = tbCustomerReference.Text
        oCmd.Parameters.Add(paramCustomerReference)

        Dim paramUpdatedBy As SqlParameter = New SqlParameter("@UpdatedBy", SqlDbType.Int)
        paramUpdatedBy.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUpdatedBy)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnRefreshNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetNotes()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("show") Then
            Call ToggleNotesGrid()
        End If
    End Sub
    
    Protected Sub gvNotes_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Call GetNotes()
    End Sub

    Protected Sub gvEvents_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Select Case pnDisplayMode
            Case 1, 2
                Call ShowEvents(bEventTypeAll:=True, bCustomerTypeAll:=rbCMAllCustomers.Checked)
            Case 3, 4
                Call ShowEvents(bEventTypeAll:=False, bCustomerTypeAll:=rbCMAllCustomers.Checked)
        End Select
    End Sub
    
    Protected Sub lnkbtnShowHideNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleNotesGrid()
    End Sub
    
    Protected Sub ToggleNotesGrid()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("hide") Then
            lnkbtnShowHideNotes.Text = "show notes"
            trNotes.Visible = False
        Else
            lnkbtnShowHideNotes.Text = "hide notes"
            trNotes.Visible = True
        End If
    End Sub
    
    Protected Sub QueueTimedEmailReminder(ByVal nCustomerKey As Integer, ByVal nReference As Integer, ByVal dtScheduledSend As DateTime, ByVal sRecipient As String, ByVal sBodyText As String, ByVal sBodyHTML As String, ByVal nQueuedBy As Integer)
        Dim bError As Boolean = False

        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToTimedQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = "CAL_MGD_REMINDER"
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = nCustomerKey
    
            oCmd.Parameters.Add(New SqlParameter("@Type", SqlDbType.Int))
            oCmd.Parameters("@Type").Value = 1
    
            oCmd.Parameters.Add(New SqlParameter("@Reference", SqlDbType.Int, 4))
            oCmd.Parameters("@Reference").Value = pnEventId
    
            oCmd.Parameters.Add(New SqlParameter("@ScheduledSend", SqlDbType.SmallDateTime))
            oCmd.Parameters("@ScheduledSend").Value = dtScheduledSend
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = "Reminder for " & lblCustomer.Text & " event " & lblEventName.Text
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int, 4))
            oCmd.Parameters("@QueuedBy").Value = nQueuedBy
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            bError = True
        Finally
            oConn.Close()
        End Try
    
        If bError Then
            WebMsgBox.Show("Unable to process request due to an internal error.")
        End If
    End Sub
    
    Protected Sub btnAddReminder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sTemp As String
        If Not IsDate(tbEmailReminderDate.Text) Then
            WebMsgBox.Show("Please enter a valid date in the format dd-mmm-yyyy, eg 20-Jun-2009")
            tbEmailReminderDate.Focus()
            Exit Sub
        End If
        
        Dim dtSendDate As DateTime = DateTime.Parse(tbEmailReminderDate.Text)
        Dim dtScheduledSend As New DateTime(dtSendDate.Year, dtSendDate.Month, dtSendDate.Day, CInt(ddlEmailReminderTime.SelectedValue.Substring(0, 2)), CInt(ddlEmailReminderTime.SelectedValue.Substring(2, 2)), 0)
        If dtScheduledSend < DateTime.Now Then
            WebMsgBox.Show("The reminder date and time you selected has already passed! Consider this message as your reminder.")
            tbEmailReminderDate.Focus()
            Exit Sub
        End If
        
        sTemp = tbEmailReminderText.Text.Trim
        If sTemp.Length = 0 Then
            WebMsgBox.Show("Please enter text for your reminder")
            tbEmailReminderText.Focus()
            Exit Sub
        End If
        sTemp = tbEmailReminderAddr.Text.Trim(", ".ToCharArray)
        If sTemp.Contains(",") Then
            Dim sEmailAddrs() As String = sTemp.Split(",".ToCharArray)
            For Each sEmailAddr As String In sEmailAddrs
                If Not bIsValidEmailAddress(sEmailAddr) Then
                    WebMsgBox.Show("Invalid email address " & sEmailAddr)
                    tbEmailReminderAddr.Focus()
                    Exit Sub
                End If
            Next
            Call AddReminder(sEmailAddrs, dtScheduledSend)
        Else
            If sTemp.Length = 0 Then
                WebMsgBox.Show("Please enter an email address")
                tbEmailReminderAddr.Focus()
                Exit Sub
            Else
                If Not bIsValidEmailAddress(sTemp) Then
                    WebMsgBox.Show("Invalid email address " & sTemp)
                    tbEmailReminderAddr.Focus()
                    Exit Sub
                Else
                    Call AddReminder(sTemp, dtScheduledSend)
                End If
            End If
        End If
        Call GetAccountHandlerEmailAddr()
        tbEmailReminderDate.Text = String.Empty
        ddlEmailReminderTime.SelectedIndex = 0
        tbEmailReminderText.Text = String.Empty
        Call GetReminders()
        WebMsgBox.Show("Reminder set")
    End Sub

    Protected Function bIsValidEmailAddress(ByRef sEmailAddr As String) As Boolean
        sEmailAddr = Trim$(sEmailAddr)
        Return Regex.IsMatch(sEmailAddr, "^[\w\.\-]+@[a-zA-Z0-9\-]+(\.[a-zA-Z0-9\-]{1,})*(\.[a-zA-Z]{2,3}){1,2}$")
    End Function

    Protected Sub AddReminder(ByVal sEmailAddr As String, ByVal dtScheduledSend As DateTime)
        Call QueueTimedEmailReminder(0, pnEventId, dtScheduledSend, sEmailAddr, tbEmailReminderText.Text, tbEmailReminderText.Text, Session("UserKey"))
    End Sub
    
    Protected Sub AddReminder(ByVal sEmailAddr() As String, ByVal dtScheduledSend As DateTime)
        For Each sAddr As String In sEmailAddr
            Call AddReminder(sAddr, dtScheduledSend)
        Next
    End Sub
    
    Protected Sub lnkbtnShowHideReminders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleRemindersGrid()
    End Sub

    Protected Sub ToggleRemindersGrid()
        If lnkbtnShowHideReminders.Text.ToLower.Contains("hide") Then
            lnkbtnShowHideReminders.Text = "show reminders"
            trReminders.Visible = False
        Else
            lnkbtnShowHideReminders.Text = "hide reminders"
            trReminders.Visible = True
        End If
    End Sub
    
    Protected Sub lnkbtrnRefreshReminders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetReminders()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("show") Then
            Call ToggleRemindersGrid()
        End If
    End Sub
    
    Protected Sub gvReminders_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.Cells.Count >= 3 Then      ' check if one or more rows - if no rows there will only be a single cell with the empty grid message
            row.Cells(1).Visible = False  ' hide items required in the query but not to be displayed (EventId, BookedBy)
        End If
    End Sub

    Protected Sub gvReminders_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Call GetReminders()
    End Sub
    
    Protected Sub lnkbtnDeleteReminder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call DeleteReminder(CInt(lb.CommandArgument))
    End Sub
    
    Protected Sub DeleteReminder(ByVal nQueueKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_DeleteReminder", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramQueueKey As SqlParameter = New SqlParameter("@QueueKey", SqlDbType.Int)
        paramQueueKey.Value = nQueueKey
        oCmd.Parameters.Add(paramQueueKey)

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.Int)
        paramUserId.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserId)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
        WebMsgBox.Show("Reminder deleted")
    End Sub
    
    Protected Sub btnPickEventProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PickEventProducts()
    End Sub
    
    Protected Sub PickEventProducts()
        Dim sSQL As String = String.Empty
        Dim bProductsAvailableToPick As Boolean = True
        Dim dictEventProducts As Dictionary(Of Integer, Integer) = dictGetEventProducts()
        For Each kv As KeyValuePair(Of Integer, Integer) In dictEventProducts
            If kv.Value = 0 Then
                WebMsgBox.Show(GetProductInfo(kv.Key) & " is not available to pick")
                bProductsAvailableToPick = False
                Exit For
            End If
        Next
        If bProductsAvailableToPick Then
            sSQL = "SELECT BookedBy FROM CalendarManagedItemEvent WHERE [id] = " & pnEventId
            Dim nBookedBy As Integer = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
            Call PickProducts(dictEventProducts, nBookedBy)
            sSQL = "UPDATE CalendarManagedItemEvent SET ProductPickDate = GETDATE() WHERE [id] = " & pnEventId
            If Not ExecuteNonQuery(sSQL) Then
                WebMsgBox.Show("Error - could not store product pick date - please inform software development.")
            End If
            'psRefreshStoredProcedure = psRefreshStoredProcedure.Replace("Unreviewed", "All")
            Call RefreshEvents()
            lblProductPickedFlag.Text = "PICKED " & Format(Now, "d-MMM-yyyy h:mm")
        Else
            Call AddNote("Tried to pick product(s) but one or more was unavailable", False)
        End If
    End Sub

    Protected Function GetProductInfo(ByVal nLogisticProductKey As Integer) As String
        GetProductInfo = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oCmd As New SqlCommand
        Dim sTemp As String = String.Empty
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.CommandText = "spASPNET_CalendarManaged_GetEventItemByKey"

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = nLogisticProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                sTemp = oDataReader("ProductDate")
                If sTemp.Length > 0 Then
                    sTemp = " (Value Date: " & sTemp & ") "
                End If
                sTemp = "Product " & oDataReader("ProductDate") & sTemp & " - " & oDataReader("ProductDescription")
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetProductInfo(): Could not retrieve data - " & ex.Message)
        Finally
            oConn.Close()
        End Try
        GetProductInfo = sTemp
    End Function
    
    Protected Function dictGetEventProducts() As Dictionary(Of Integer, Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oCmd As New SqlCommand
        Dim dictProducts As New Dictionary(Of Integer, Integer)
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.CommandText = "spASPNET_CalendarManaged_GetEventItemKeysById"

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read()
                    dictProducts.Add(oDataReader("LogisticProductKey"), oDataReader("Quantity"))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("dictEventProducts(): Could not retrieve data - " & ex.Message)
        Finally
            oConn.Close()
        End Try
        dictGetEventProducts = dictProducts
    End Function
    
    Protected Function GetAccountHandlerDetails() As String
        GetAccountHandlerDetails = String.Empty
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oCmd As New SqlCommand
        Dim sTemp As String = String.Empty
        oCmd.Connection = oConn
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.CommandText = "spASPNET_CalendarManaged_GetAccountHandlerForCustomer"

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = hidCustomerKey.Value
        oCmd.Parameters.Add(paramCustomerKey)
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                sTemp = oDataReader("Name") & "~" & oDataReader("EmailAddr")
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetAccountHandlerDetails(): Could not retrieve data - " & ex.Message)
        Finally
            oConn.Close()
        End Try
        GetAccountHandlerDetails = sTemp
    End Function
    
    Protected Function GetEventSummaryForSpecialInstructions() As String
        GetEventSummaryForSpecialInstructions = String.Empty
        Dim sbEventDetails As New StringBuilder
        sbEventDetails.Append(" (")
        sbEventDetails.Append(" EVNT: ")
        sbEventDetails.Append(lblEventName.Text)
        sbEventDetails.Append(" DELIV: ")
        sbEventDetails.Append(lblDeliveryDate.Text)
        sbEventDetails.Append(" TIME: ")
        If ddlDeliveryTime.SelectedItem.Text.Contains("special") Then
            sbEventDetails.Append("(see Spcl Instrs)")
        Else
            sbEventDetails.Append(ddlDeliveryTime.SelectedItem.Text)
        End If
        sbEventDetails.Append(" TO ")
        sbEventDetails.Append(tbPreciseDeliveryPoint.Text)
        sbEventDetails.Append(" COLLECT: ")
        sbEventDetails.Append(lblCollectionDate.Text)
        sbEventDetails.Append(" TIME: ")
        If ddlCollectionTime.SelectedItem.Text.Contains("special") Then
            sbEventDetails.Append("(see Spcl Instrs)")
        Else
            sbEventDetails.Append(ddlCollectionTime.SelectedItem.Text)
        End If
        sbEventDetails.Append(" FROM: ")
        sbEventDetails.Append(tbPreciseCollectionPoint.Text)
        
        If tbContactPhone.Text.Trim <> String.Empty Then
            sbEventDetails.Append(" CTC PHONE: ")
            sbEventDetails.Append(tbContactPhone.Text.Trim)
        End If
        If tbContactMobile.Text.Trim <> String.Empty Then
            sbEventDetails.Append(" CTC MOB: ")
            sbEventDetails.Append(tbContactMobile.Text.Trim)
        End If
        If tbCMContactName2.Text.Trim <> String.Empty Then
            sbEventDetails.Append(" CTC 2: ")
            sbEventDetails.Append(tbCMContactName2.Text.Trim)
        End If
        If tbCMContactPhone2.Text.Trim <> String.Empty Then
            sbEventDetails.Append(" CTC PHONE 2: ")
            sbEventDetails.Append(tbCMContactPhone2.Text.Trim)
        End If
        If tbCMContactMobile2.Text.Trim <> String.Empty Then
            sbEventDetails.Append(" CTC MOB 2: ")
            sbEventDetails.Append(tbCMContactMobile2.Text.Trim)
        End If
        
        If cbDifferentCollectionAddress.Checked Then
            tbCollectionAddress1.Text = tbCollectionAddress1.Text.Trim()
            tbCollectionAddress2.Text = tbCollectionAddress2.Text.Trim()
            tbCollectionTown.Text = tbCollectionTown.Text.Trim()
            tbCollectionPostcode.Text = tbCollectionPostcode.Text.Trim()
            
            If Not (tbCollectionAddress1.Text = String.Empty And tbCollectionAddress2.Text = String.Empty And tbCollectionTown.Text = String.Empty And tbCollectionPostcode.Text = String.Empty) Then
                If Not tbCollectionAddress1.Text = String.Empty Then
                    sbEventDetails.Append(" COLLECT ADDR1: ")
                    sbEventDetails.Append(tbCollectionAddress1.Text)
                End If
                If Not tbCollectionAddress2.Text = String.Empty Then
                    sbEventDetails.Append(" COLLECT ADDR2: ")
                    sbEventDetails.Append(tbCollectionAddress2.Text)
                End If
                If Not tbCollectionTown.Text = String.Empty Then
                    sbEventDetails.Append(" COLLECT TOWN: ")
                    sbEventDetails.Append(tbCollectionTown.Text)
                End If
                If Not tbCollectionPostcode.Text = String.Empty Then
                    sbEventDetails.Append(" COLLECT POSTCODE: ")
                    sbEventDetails.Append(tbCollectionPostcode.Text)
                End If
                If trCMCollectionCountry.Visible Then
                    If Not ddlCMCollectionCountry.SelectedValue = COUNTRY_KEY_UK Then
                        sbEventDetails.Append(" COLLECT CTRY: ")
                        sbEventDetails.Append(ddlCMCollectionCountry.SelectedItem.Text)
                    End If
                End If
            End If
        End If
        sbEventDetails.Append(")")
        GetEventSummaryForSpecialInstructions = sbEventDetails.ToString
    End Function
    
    Protected Sub PickProducts(ByVal dictProducts As Dictionary(Of Integer, Integer), ByVal nBookedBy As Integer)
        Dim sSQL As String
        Dim oDataTable As DataTable
        Dim sAH As String = GetAccountHandlerDetails()
        Dim sAccountHandlerDetails() As String
        Dim sAccountHandlerEmailAddr As String = String.Empty
        If sAH.Contains("~") Then
            sAccountHandlerDetails = sAH.Split("~".ToCharArray)
        End If

        Dim nCustomerKey As Integer = 5
        For Each kv As KeyValuePair(Of Integer, Integer) In dictProducts
            sSQL = "SELECT CustomerKey FROM LogisticProduct WHERE LogisticProductKey = " & kv.Key
            oDataTable = ExecuteQueryToDataTable(sSQL)
            nCustomerKey = oDataTable.Rows(0).Item(0)
            Exit For
        Next
        sSQL = "SELECT ISNULL(CustomerName,''), ISNULL(CustomerAddr1,''), ISNULL(CustomerAddr2,''), ISNULL(CustomerAddr3,''), ISNULL(CustomerTown,''), ISNULL(CustomerCounty,''), ISNULL(CustomerPostCode,''), ISNULL(CustomerCountryKey,0) FROM Customer WHERE CustomerKey = " & nCustomerKey
        oDataTable = ExecuteQueryToDataTable(sSQL)
        Dim oDataRow As DataRow = oDataTable.Rows(0)
        
        Dim sSpecialInstr As String
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        'param1.Value = CLng(Session("UserKey"))
        param1.Value = nBookedBy
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = hidCustomerKey.Value
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param3)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param4)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = lblEventName.Text
        oCmdAddBooking.Parameters.Add(param5)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = tbCustomerReference.Text
        oCmdAddBooking.Parameters.Add(param6)
            
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
            
            
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        sSpecialInstr = tbSpecialInstructions.Text
        sSpecialInstr = Replace(sSpecialInstr, vbCrLf, " ")
        If sSpecialInstr <> String.Empty Then
            sSpecialInstr += " "
        End If
        sSpecialInstr += "SYSTEM: SEND VIA APC!! " & GetEventSummaryForSpecialInstructions()
        'sSpecialInstr += GetEventSummaryForSpecialInstructions()
        param7.Value = sSpecialInstr
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param8)

        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param9)

        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)

        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = "PRINTED MATTER - FREE DOMICILE"
        oCmdAddBooking.Parameters.Add(param11)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = oDataRow(0)
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = oDataRow(1)
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = oDataRow(2)
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = oDataRow(3)
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = oDataRow(4)
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = oDataRow(5)
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = oDataRow(6)
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        'param20.Value = COUNTRY_KEY_UK
        param20.Value = oDataRow(7)
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        If sAH.Length > 0 Then
            param21.Value = sAccountHandlerDetails(0)
        Else
            param21.Value = "Operations Manager"
        End If
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = "020 8751 1111"
        oCmdAddBooking.Parameters.Add(param22)
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        If sAH.Length > 0 Then
            param23.Value = sAccountHandlerDetails(1)
        Else
            param23.Value = "account.managers@transworld.eu.com"
        End If
        oCmdAddBooking.Parameters.Add(param23)
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = tbContactName.Text
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = tbEventAddress1.Text
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = tbEventAddress2.Text
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = tbEventAddress3.Text
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = tbTown.Text
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = tbPostcode.Text
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        If trCMCountry.Visible Then
            param32.Value = ddlCMCountry.SelectedValue
        Else
            param32.Value = COUNTRY_KEY_UK
        End If
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = tbContactName.Text
        oCmdAddBooking.Parameters.Add(param33)
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = tbContactPhone.Text
        oCmdAddBooking.Parameters.Add(param34)
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param35)
        Dim param36 As SqlParameter = New SqlParameter("@CneePreAlertFlag", SqlDbType.Bit)
        param36.Value = 0
        oCmdAddBooking.Parameters.Add(param36)
        Dim param37 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)
        Dim param38 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param38.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param38)
        Try
            BookingFailed = False
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
            If lBookingKey > 0 Then
                If dictProducts.Count > 0 Then
                    For Each kv As KeyValuePair(Of Integer, Integer) In dictProducts
                        Dim lProductKey As Long = kv.Key
                        Dim lPickQuantity As Long = 1
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                        param51.Value = CLng(Session("UserKey"))
                        oCmdAddStockItem.Parameters.Add(param51)
                        Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                        param52.Value = CLng(hidCustomerKey.Value)
                        oCmdAddStockItem.Parameters.Add(param52)
                        Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                        param53.Value = lBookingKey
                        oCmdAddStockItem.Parameters.Add(param53)
                        Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param54.Value = lProductKey
                        oCmdAddStockItem.Parameters.Add(param54)
                        Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                        param55.Value = "PENDING"
                        oCmdAddStockItem.Parameters.Add(param55)
                        Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                        param56.Value = lPickQuantity
                        oCmdAddStockItem.Parameters.Add(param56)
                        Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                        param57.Value = lConsignmentKey
                        oCmdAddStockItem.Parameters.Add(param57)
                        oCmdAddStockItem.Connection = oConn
                        oCmdAddStockItem.Transaction = oTrans
                        oCmdAddStockItem.ExecuteNonQuery()
                    Next
                    Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                    oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                    Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param71.Value = lBookingKey
                    oCmdCompleteBooking.Parameters.Add(param71)
                    oCmdCompleteBooking.Connection = oConn
                    oCmdCompleteBooking.Transaction = oTrans
                    oCmdCompleteBooking.ExecuteNonQuery()

                Else
                    BookingFailed = True
                    WebMsgBox.Show("No stock items found for booking")
                End If
            Else
                BookingFailed = True
                WebMsgBox.Show("Error adding Web Booking [BookingKey=0].")
                Call AddNote("Error during product pick [BookingKey=0]", False)
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                WebMsgBox.Show("Consignment booking successful - consignment number is " & lConsignmentKey.ToString)
                Dim sPlural As String = String.Empty
                If dictProducts.Count > 1 Then
                    sPlural = "s"
                End If
                Call AddNote("Product" & sPlural & " picked in consignment " & lConsignmentKey.ToString, True)
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            WebMsgBox.Show("PickProducts: " & ex.ToString)
            oTrans.Rollback("AddBooking")
            Call AddNote("Error during product pick: " & ex.Message, False)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ddlSiteAdministrator_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nOldUserKey As Integer
        If IsNumeric(hidSiteAdministratorKey.Value) Then
            nOldUserKey = hidSiteAdministratorKey.Value
        Else
            nOldUserKey = 0
        End If
        Call SetNewSiteAdministrator(nOldUserKey, ddlSiteAdministrator.SelectedValue)
        hidSiteAdministratorKey.Value = ddlSiteAdministrator.SelectedValue
    End Sub

    Protected Sub SetNewSiteAdministrator(ByVal nOldUserKey As Integer, ByVal nNewUserKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetNewSiteAdministrator2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramOldUserKey As SqlParameter = New SqlParameter("@OldUserKey", SqlDbType.Int)
        paramOldUserKey.Value = nOldUserKey
        oCmd.Parameters.Add(paramOldUserKey)

        Dim paramNewUserKey As SqlParameter = New SqlParameter("@NewUserKey", SqlDbType.Int)
        paramNewUserKey.Value = nNewUserKey
        oCmd.Parameters.Add(paramNewUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSaveSiteFeatureChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveSiteFeatureChanges()
        If cbFileUpload.Checked Then
            Call SaveFileUploadNotificationEmailAddresses()
        End If
    End Sub

    Protected Sub SaveSiteFeatureChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_SiteContent3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramAction As New SqlParameter("@Action", SqlDbType.NVarChar, 50)
        paramAction.Value = "SET"
        oCmd.Parameters.Add(paramAction)

        Dim paramSiteKey As New SqlParameter("@SiteKey", SqlDbType.Int)
        paramSiteKey.Value = Session("SiteKey")
        oCmd.Parameters.Add(paramSiteKey)

        Dim paramContentType As New SqlParameter("@ContentType", SqlDbType.NVarChar, 50)
        paramContentType.Value = "ProtectedSiteSettings4"
        oCmd.Parameters.Add(paramContentType)

        Dim paramDefaultDescription As New SqlParameter("@DefaultDescription", SqlDbType.NVarChar, 50)
        paramDefaultDescription.Value = tbDefaultDescription.Text
        oCmd.Parameters.Add(paramDefaultDescription)

        Dim paramPostcodeLookup As New SqlParameter("@PostcodeLookup", SqlDbType.Bit)
        paramPostcodeLookup.Value = cbPostcodeLookup.Checked
        oCmd.Parameters.Add(paramPostcodeLookup)

        Dim paramCalendarManagement As New SqlParameter("@CalendarManagement", SqlDbType.Bit)
        paramCalendarManagement.Value = cbCalendarManagement.Checked
        oCmd.Parameters.Add(paramCalendarManagement)

        Dim paramUserPermissions As New SqlParameter("@UserPermissions", SqlDbType.Bit)
        paramUserPermissions.Value = cbUserPermissions.Checked
        oCmd.Parameters.Add(paramUserPermissions)

        Dim paramFileUpload As New SqlParameter("@FileUpload", SqlDbType.Bit)
        paramFileUpload.Value = cbFileUpload.Checked
        oCmd.Parameters.Add(paramFileUpload)

        Dim paramUseLabelPrinter As New SqlParameter("@UseLabelPrinter", SqlDbType.Bit)
        paramUseLabelPrinter.Value = cbUseLabelPrinter.Checked
        oCmd.Parameters.Add(paramUseLabelPrinter)

        Dim paramSearchCompanyNameOnly As New SqlParameter("@SearchCompanyNameOnly", SqlDbType.Bit)
        paramSearchCompanyNameOnly.Value = cbSearchCompanyNameOnly.Checked
        oCmd.Parameters.Add(paramSearchCompanyNameOnly)

        Dim paramMakeRef1Mandatory As New SqlParameter("@MakeRef1Mandatory", SqlDbType.Bit)
        paramMakeRef1Mandatory.Value = cbMakeRef1Mandatory.Checked
        oCmd.Parameters.Add(paramMakeRef1Mandatory)

        Dim paramRef1Label As New SqlParameter("@Ref1Label", SqlDbType.NVarChar, 50)
        paramRef1Label.Value = tbRef1Label.Text
        oCmd.Parameters.Add(paramRef1Label)

        Dim paramMakeRef2Mandatory As New SqlParameter("@MakeRef2Mandatory", SqlDbType.Bit)
        paramMakeRef2Mandatory.Value = cbMakeRef2Mandatory.Checked
        oCmd.Parameters.Add(paramMakeRef2Mandatory)

        Dim paramRef2Label As New SqlParameter("@Ref2Label", SqlDbType.NVarChar, 50)
        paramRef2Label.Value = tbRef2Label.Text
        oCmd.Parameters.Add(paramRef2Label)

        Dim paramMakeRef3Mandatory As New SqlParameter("@MakeRef3Mandatory", SqlDbType.Bit)
        paramMakeRef3Mandatory.Value = cbMakeRef3Mandatory.Checked
        oCmd.Parameters.Add(paramMakeRef3Mandatory)

        Dim paramRef3Label As New SqlParameter("@Ref3Label", SqlDbType.NVarChar, 50)
        paramRef3Label.Value = tbRef3Label.Text
        oCmd.Parameters.Add(paramRef3Label)

        Dim paramMakeRef4Mandatory As New SqlParameter("@MakeRef4Mandatory", SqlDbType.Bit)
        paramMakeRef4Mandatory.Value = cbMakeRef4Mandatory.Checked
        oCmd.Parameters.Add(paramMakeRef4Mandatory)

        Dim paramRef4Label As New SqlParameter("@Ref4Label", SqlDbType.NVarChar, 50)
        paramRef4Label.Value = tbRef4Label.Text
        oCmd.Parameters.Add(paramRef4Label)

        Dim paramThirdPartyCollectionKey As New SqlParameter("@ThirdPartyCollectionKey", SqlDbType.Int)
        paramThirdPartyCollectionKey.Value = CInt(tbThirdPartyCollectionKey.Text)
        oCmd.Parameters.Add(paramThirdPartyCollectionKey)

        Dim paramHideCollectionButton As New SqlParameter("@HideCollectionButton", SqlDbType.Bit)
        paramHideCollectionButton.Value = cbHideCollectionButton.Checked
        oCmd.Parameters.Add(paramHideCollectionButton)

        Dim paramProductOwners As New SqlParameter("@ProductOwners", SqlDbType.Bit)
        paramProductOwners.Value = cbProductOwners.Checked
        oCmd.Parameters.Add(paramProductOwners)

        Dim paramCategoryCount As New SqlParameter("@CategoryCount", SqlDbType.TinyInt)
        paramCategoryCount.Value = CInt(tbCategoryCount.Text)
        oCmd.Parameters.Add(paramCategoryCount)

        Dim paramMisc1 As New SqlParameter("@Misc1", SqlDbType.Bit)
        paramMisc1.Value = cbPrintOndemandTab.Checked
        oCmd.Parameters.Add(paramMisc1)

        'Dim paramMisc2 As New SqlParameter("@Misc2", SqlDbType.Bit)
        'paramMisc2.Value = cbOnDemandProducts.Checked
        'oCmd.Parameters.Add(paramMisc2)

        Dim paramMisc2 As New SqlParameter("@Misc2", SqlDbType.Bit)
        paramMisc2.Value = False
        oCmd.Parameters.Add(paramMisc2)

        Dim paramMisc3 As New SqlParameter("@Misc3", SqlDbType.Bit)
        paramMisc3.Value = False
        oCmd.Parameters.Add(paramMisc3)

        Dim paramMisc4 As New SqlParameter("@Misc4", SqlDbType.Bit)
        paramMisc4.Value = False
        oCmd.Parameters.Add(paramMisc4)

        Dim paramMisc5 As New SqlParameter("@Misc5", SqlDbType.Bit)
        paramMisc5.Value = False
        oCmd.Parameters.Add(paramMisc5)

        Dim paramMisc6 As New SqlParameter("@Misc6", SqlDbType.Bit)
        paramMisc6.Value = False
        oCmd.Parameters.Add(paramMisc6)

        Dim paramOnDemandProducts As New SqlParameter("@OnDemandProducts", SqlDbType.Bit)
        paramOnDemandProducts.Value = cbOnDemandProducts.Checked
        oCmd.Parameters.Add(paramOnDemandProducts)

        Dim paramCustomLetters As New SqlParameter("@CustomLetters", SqlDbType.Bit)
        paramCustomLetters.Value = cbCustomLetters.Checked
        oCmd.Parameters.Add(paramCustomLetters)

        Dim paramWebForm As New SqlParameter("@WebForm", SqlDbType.Bit)
        paramWebForm.Value = cbWebForm.Checked
        oCmd.Parameters.Add(paramWebForm)

        Dim paramStockOrderCustRef1Visible As New SqlParameter("@StockOrderCustRef1Visible", SqlDbType.Bit)
        paramStockOrderCustRef1Visible.Value = cbCustRef1IsVisible.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef1Visible)

        Dim paramStockOrderCustRef1Mandatory As New SqlParameter("@StockOrderCustRef1Mandatory", SqlDbType.Bit)
        paramStockOrderCustRef1Mandatory.Value = cbCustRef1IsMandatory.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef1Mandatory)

        Dim paramStockOrderCustRefLabel1Legend As New SqlParameter("@StockOrderCustRefLabel1Legend", SqlDbType.NVarChar, 50)
        paramStockOrderCustRefLabel1Legend.Value = tbCustRef1Label.Text
        oCmd.Parameters.Add(paramStockOrderCustRefLabel1Legend)
        
        Dim paramStockOrderCustRef2Visible As New SqlParameter("@StockOrderCustRef2Visible", SqlDbType.Bit)
        paramStockOrderCustRef2Visible.Value = cbCustRef2IsVisible.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef2Visible)

        Dim paramStockOrderCustRef2Mandatory As New SqlParameter("@StockOrderCustRef2Mandatory", SqlDbType.Bit)
        paramStockOrderCustRef2Mandatory.Value = cbCustRef2IsMandatory.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef2Mandatory)

        Dim paramStockOrderCustRefLabel2Legend As New SqlParameter("@StockOrderCustRefLabel2Legend", SqlDbType.NVarChar, 50)
        paramStockOrderCustRefLabel2Legend.Value = tbCustRef2Label.Text
        oCmd.Parameters.Add(paramStockOrderCustRefLabel2Legend)
        
        Dim paramStockOrderCustRef3Visible As New SqlParameter("@StockOrderCustRef3Visible", SqlDbType.Bit)
        paramStockOrderCustRef3Visible.Value = cbCustRef3IsVisible.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef3Visible)

        Dim paramStockOrderCustRef3Mandatory As New SqlParameter("@StockOrderCustRef3Mandatory", SqlDbType.Bit)
        paramStockOrderCustRef3Mandatory.Value = cbCustRef3IsMandatory.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef3Mandatory)

        Dim paramStockOrderCustRefLabel3Legend As New SqlParameter("@StockOrderCustRefLabel3Legend", SqlDbType.NVarChar, 50)
        paramStockOrderCustRefLabel3Legend.Value = tbCustRef3Label.Text
        oCmd.Parameters.Add(paramStockOrderCustRefLabel3Legend)
        
        Dim paramStockOrderCustRef4Visible As New SqlParameter("@StockOrderCustRef4Visible", SqlDbType.Bit)
        paramStockOrderCustRef4Visible.Value = cbCustRef4IsVisible.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef4Visible)

        Dim paramStockOrderCustRef4Mandatory As New SqlParameter("@StockOrderCustRef4Mandatory", SqlDbType.Bit)
        paramStockOrderCustRef4Mandatory.Value = cbCustRef4IsMandatory.Checked
        oCmd.Parameters.Add(paramStockOrderCustRef4Mandatory)

        Dim paramStockOrderCustRefLabel4Legend As New SqlParameter("@StockOrderCustRefLabel4Legend", SqlDbType.NVarChar, 50)
        paramStockOrderCustRefLabel4Legend.Value = tbCustRef4Label.Text
        oCmd.Parameters.Add(paramStockOrderCustRefLabel4Legend)
        
        Dim paramSessionTimeout As New SqlParameter("@SessionTimeout", SqlDbType.Int)
        paramSessionTimeout.Value = CInt(tbSessionTimeout.Text)
        oCmd.Parameters.Add(paramSessionTimeout)

        Dim paramWebFormCustomerKey As New SqlParameter("@WebFormCustomerKey", SqlDbType.Int)
        'paramWebFormCustomerKey.Value = CInt(tbWebFormCustomerKey.Text)
        paramWebFormCustomerKey.Value = CInt(ddlWebFormCustomer.SelectedValue)
        oCmd.Parameters.Add(paramWebFormCustomerKey)

        Dim paramWebFormGenericUser As New SqlParameter("@WebFormGenericUserKey", SqlDbType.Int)
        'paramWebFormGenericUser.Value = CInt(tbWebFormGenericUserKey.Text)
        If ddlWebFormGenericUser.Enabled Then
            paramWebFormGenericUser.Value = CInt(ddlWebFormGenericUser.SelectedValue)
        Else
            paramWebFormGenericUser.Value = 0
        End If
        oCmd.Parameters.Add(paramWebFormGenericUser)

        Dim paramWebFormPageTitle As New SqlParameter("@WebFormPageTitle", SqlDbType.VarChar, 50)
        paramWebFormPageTitle.Value = tbWebFormPageTitle.Text
        oCmd.Parameters.Add(paramWebFormPageTitle)
        
        Dim paramWebFormLogoImage As New SqlParameter("@WebFormLogoImage", SqlDbType.VarChar, 100)
        paramWebFormLogoImage.Value = tbWebFormLogoImage.Text
        oCmd.Parameters.Add(paramWebFormLogoImage)
        
        Dim paramWebFormTopLegend As New SqlParameter("@WebFormTopLegend", SqlDbType.NVarChar, 50)
        paramWebFormTopLegend.Value = tbWebFormTopLegend.Text
        oCmd.Parameters.Add(paramWebFormTopLegend)
        
        Dim paramWebFormBottomLegend As New SqlParameter("@WebFormBottomLegend", SqlDbType.NVarChar, 50)
        paramWebFormBottomLegend.Value = tbWebFormBottomLegend.Text
        oCmd.Parameters.Add(paramWebFormBottomLegend)
        
        Dim paramWebFormShowPrice As New SqlParameter("@WebFormShowPrice", SqlDbType.Bit)
        paramWebFormShowPrice.Value = cbWebFormShowPrice.Checked
        oCmd.Parameters.Add(paramWebFormShowPrice)

        Dim paramWebFormShowZeroQuantity As New SqlParameter("@WebFormShowZeroQuantity", SqlDbType.Bit)
        paramWebFormShowZeroQuantity.Value = cbWebFormShowZeroQuantity.Checked
        oCmd.Parameters.Add(paramWebFormShowZeroQuantity)

        Dim paramWebFormZeroStockNotification As New SqlParameter("@WebFormZeroStockNotification", SqlDbType.Bit)
        paramWebFormZeroStockNotification.Value = cbWebFormZeroStockNotification.Checked
        oCmd.Parameters.Add(paramWebFormZeroStockNotification)

        Dim paramWebFormHomePageText As New SqlParameter("@WebFormHomePageText", SqlDbType.NVarChar, 1000)
        paramWebFormHomePageText.Value = FCKedWebFormHomePage.Value
        oCmd.Parameters.Add(paramWebFormHomePageText)
        
        Dim paramWebFormAddressPageText As New SqlParameter("@WebFormAddressPageText", SqlDbType.NVarChar, 1000)
        paramWebFormAddressPageText.Value = FCKedWebFormAddressPage.Value
        oCmd.Parameters.Add(paramWebFormAddressPageText)
        
        Dim paramWebFormHelpPageText As New SqlParameter("@WebFormHelpPageText", SqlDbType.NVarChar, 1000)
        paramWebFormHelpPageText.Value = FCKedWebFormHelpPage.Value
        oCmd.Parameters.Add(paramWebFormHelpPageText)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("SaveSiteFeatureChanges: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        If cbOnDemandProducts.Checked Then
            Call SaveAttemAccessParameters()
        End If
    End Sub
    
    Protected Sub SaveAttemAccessParameters()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_OnDemand_AccessSetParameters", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oCmd.Parameters.Add(New SqlParameter("@UserName", SqlDbType.VarChar, 50))
        oCmd.Parameters("@UserName").Value = tbAttemUserName.Text
        oCmd.Parameters.Add(New SqlParameter("@Password", SqlDbType.VarChar, 50))
        oCmd.Parameters("@Password").Value = tbAttemPassword.Text
        oCmd.Parameters.Add(New SqlParameter("@Customer", SqlDbType.VarChar, 50))
        oCmd.Parameters("@Customer").Value = tbAttemCustomerName.Text
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in SaveAttemAccessParameters: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SetWebFormControlsVisibility(ByVal bVisibility As Boolean)
        trWebForm01.Visible = bVisibility
        trWebForm02.Visible = bVisibility
        trWebForm03.Visible = bVisibility
        trWebForm04.Visible = bVisibility
        trWebForm05.Visible = bVisibility
        trWebForm06.Visible = bVisibility
        trWebForm07.Visible = bVisibility
        trWebForm08.Visible = bVisibility
    End Sub
    
    Protected Sub cbWebForm_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            SetWebFormControlsVisibility(True)
        Else
            SetWebFormControlsVisibility(False)
        End If
    End Sub
    
    Protected Sub cbCustRef1IsVisible_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCustRef1Enabled(cb.Checked)
    End Sub
    
    Protected Sub SetCustRef1Enabled(ByVal bEnabled As Boolean)
        cbCustRef1IsMandatory.Enabled = bEnabled
        tbCustRef1Label.Enabled = bEnabled
    End Sub

    Protected Sub cbCustRef2IsVisible_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCustRef2Enabled(cb.Checked)
    End Sub

    Protected Sub SetCustRef2Enabled(ByVal bEnabled As Boolean)
        cbCustRef2IsMandatory.Enabled = bEnabled
        tbCustRef2Label.Enabled = bEnabled
    End Sub

    Protected Sub cbCustRef3IsVisible_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCustRef3Enabled(cb.Checked)
    End Sub

    Protected Sub SetCustRef3Enabled(ByVal bEnabled As Boolean)
        cbCustRef3IsMandatory.Enabled = bEnabled
        tbCustRef3Label.Enabled = bEnabled
    End Sub

    Protected Sub cbCustRef4IsVisible_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCustRef4Enabled(cb.Checked)
    End Sub
    
    Protected Sub SetCustRef4Enabled(ByVal bEnabled As Boolean)
        cbCustRef4IsMandatory.Enabled = bEnabled
        tbCustRef4Label.Enabled = bEnabled
    End Sub

    Protected Sub cbOnDemandProducts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            trAttemAccess.Visible = True
        Else
            trAttemAccess.Visible = False
        End If
    End Sub
    
    Protected Sub ddlWebFormCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            ddlWebFormGenericUser.Enabled = True
            Call SetWebFormGenericUser(CInt(ddlWebFormCustomer.SelectedValue), 0)
        Else
            ddlWebFormGenericUser.Enabled = False
        End If
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
            WebMsgBox.Show("Error in ExecuteNonQuery executing " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub cbFileUpload_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            lblLegendFileUploadNotificationEmailAddresses.Visible = True
            tbFileUploadNotificationEmailAddresses.Visible = True
            tbFileUploadNotificationEmailAddresses.Text = String.Empty
        Else
            lblLegendFileUploadNotificationEmailAddresses.Visible = False
            tbFileUploadNotificationEmailAddresses.Visible = False
            Call ExecuteNonQuery("DELETE FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey"))
        End If
    End Sub
    
    Protected Sub SaveFileUploadNotificationEmailAddresses()
        tbFileUploadNotificationEmailAddresses.Text = tbFileUploadNotificationEmailAddresses.Text.Trim
        If tbFileUploadNotificationEmailAddresses.Text = String.Empty Then
            Call ExecuteNonQuery("DELETE FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey"))
        Else
            Dim EmailAddrs() As String = tbFileUploadNotificationEmailAddresses.Text.Split(",")
            Dim regexEmailAddress = "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
            Dim bValid As Boolean = True
            For Each EmailAddr As String In EmailAddrs
                If Not Regex.IsMatch(EmailAddr, regexEmailAddress) Then
                    WebMsgBox.Show(EmailAddr & " is not a recognised valid email address for file upload notifications. Please correct.")
                    bValid = False
                End If
            Next
            If bValid Then
                Call ExecuteNonQuery("DELETE FROM FileUploadNotification WHERE CustomerKey = " & Session("CustomerKey"))
                For Each EmailAddr As String In EmailAddrs
                    If EmailAddr.Trim <> String.Empty Then
                        Call ExecuteNonQuery("INSERT INTO FileUploadNotification (EmailAddr, CustomerKey) VALUES ('" & EmailAddr.Trim.Replace("'", "''") & "', " & Session("CustomerKey") & ")")
                    End If
                Next
            End If
        End If
    End Sub

    Protected Sub cbDifferentCollectionAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCollectionAddressVisibility(cb.Checked)
    End Sub
    
    Protected Sub SetCollectionAddressVisibility(ByVal bVisibility As Boolean)
        trCollectionAddress1.Visible = bVisibility
        trCollectionAddress2.Visible = bVisibility
        If Not bVisibility Then
            tbCollectionAddress1.Text = String.Empty
            tbCollectionAddress2.Text = String.Empty
            tbCollectionTown.Text = String.Empty
            tbCollectionPostcode.Text = String.Empty
        End If
    End Sub

    Protected Sub lnkbtnEmailAddrMe_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT EmailAddr FROM UserProfile WHERE [key] = " & Session("UserKey"))
        If oDataTable.Rows.Count > 0 Then
            tbEmailReminderAddr.Text = oDataTable.Rows(0).Item(0)
            tbEmailReminderText.Focus()
        End If
    End Sub

    Property pnEventId() As Integer
        Get
            Dim o As Object = ViewState("AH_EventId")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("AH_EventId") = Value
        End Set
    End Property
    
    Property pnDisplayMode() As Integer
        Get
            Dim o As Object = ViewState("AH_DisplayMode")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("AH_DisplayMode") = Value
        End Set
    End Property
    
    Property psRefreshStoredProcedure() As String
        Get
            Dim o As Object = ViewState("AH_RefreshStoredProcedure")
            If o Is Nothing Then
                Return "NOT DEFINED"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AH_RefreshStoredProcedure") = Value
        End Set
    End Property

    Protected Sub lnkbtnCMAddSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetContact2FieldsVisibility(True)
        tbCMContactName2.Focus()
    End Sub

    Protected Sub lnkbtnCMRemoveSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbCMContactName2.Text = String.Empty
        tbCMContactPhone2.Text = String.Empty
        tbCMContactMobile2.Text = String.Empty
        Call SetContact2FieldsVisibility(False)
    End Sub

    Protected Sub lnkbtnCMAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCountry.Visible = True
        ddlCMCountry.SelectedIndex = 0
        ddlCMCountry.Focus()
    End Sub

    Protected Sub lnkbtnCMCollectionAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCollectionCountry.Visible = True
        ddlCMCollectionCountry.SelectedIndex = 0
        ddlCMCollectionCountry.Focus()
    End Sub
    
    Protected Sub SetContact2FieldsVisibility(ByVal bVisibility As Boolean)
        lnkbtnCMAddSecondContact.Visible = Not bVisibility
        rfvCMContactName2.Visible = bVisibility
        lblLegendCMContactName2.Visible = bVisibility
        tbCMContactName2.Visible = bVisibility
        rfvCMContactMobile2.Visible = bVisibility
        lblLegendCMContactMobile2.Visible = bVisibility
        tbCMContactMobile2.Visible = bVisibility
        trCMContactPhone2.Visible = bVisibility
    End Sub
    
    Protected Sub InitCountryDropdowns()
        If ddlCMCountry.Items.Count = 0 Or ddlCMCollectionCountry.Items.Count = 0 Then
            Dim sSQL As String = "SELECT SUBSTRING(CountryName,1,25) 'CountryName', CountryKey FROM Country WHERE DeletedFlag = 0 ORDER BY CountryName"
            Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CountryName", "CountryKey")
            ddlCMCountry.Items.Clear()
            ddlCMCollectionCountry.Items.Clear()
            ddlCMCountry.Items.Add(New ListItem("- please select -", 0))
            ddlCMCollectionCountry.Items.Add(New ListItem("- please select -", 0))
            For Each li As ListItem In oListItemCollection
                ddlCMCountry.Items.Add(li)
                ddlCMCollectionCountry.Items.Add(li)
            Next
        End If
    End Sub
    
    Protected Sub cbIncludeCancelledEvents_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cb.Font.Bold = True
            cb.ForeColor = Drawing.Color.Red
        Else
            cb.Font.Bold = False
            cb.ForeColor = Drawing.Color.Empty
        End If
        gvEvents.Visible = False
        pnlEvent.Visible = False
    End Sub
    
    Protected Sub ddlCMRetrospectiveDays_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        gvEvents.PageIndex = 0
        Call RefreshEvents()
    End Sub

    Protected Sub ddlCMItemsPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        gvEvents.PageIndex = 0
        Call RefreshEvents()
    End Sub
    
    Protected Sub cbRetrospective_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        ddlCMRetrospectiveDays.Enabled = cb.Checked
        If cb.Checked Then
            gvEvents.PageIndex = 0
            Call RefreshEvents()
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Account Handler</title>
</head>
<body>
    <form id="form1" runat="server">
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <strong style="color: navy; font-size:x-small; font-family:Verdana">&nbsp;Account Handler<br />
            </strong>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                        <asp:LinkButton ID="lnkbtnCMHelp" runat="server" OnClientClick='window.open("help_cmproductsah.pdf", "CMHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");' Font-Names="Verdana" Font-Size="XX-Small">calendar managed products help (acct handler)</asp:LinkButton></td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:Button ID="btnShowEventsAwaitingReview" CommandArgument="awaitingreview" Text="show events awaiting review" OnClick="btnShowEvents_Click" runat="server" />&nbsp;
                        <asp:Button ID="btnShowAllEvents" CommandArgument="all" Text="show all events" OnClick="btnShowEvents_Click" runat="server" />
                        &nbsp;&nbsp;<asp:RadioButton ID="rbCMAllCustomers" runat="server" GroupName="CustomerType" Text="all customers" AutoPostBack="True" OnCheckedChanged="rbCMAllCustomers_CheckedChanged" Checked="True" Font-Names="Verdana" Font-Size="XX-Small" />
                        <asp:RadioButton ID="rbCMMyCustomers" runat="server" GroupName="CustomerType" Text="my customers" AutoPostBack="True" OnCheckedChanged="rbCMMyCustomers_CheckedChanged" Font-Names="Verdana" Font-Size="XX-Small" />
                        &nbsp; &nbsp; &nbsp;
                        <asp:Label ID="lblLegendSelectAccountHandler" runat="server" 
                            Text="account handler:" Visible="False" Font-Names="Verdana" 
                            Font-Size="XX-Small"></asp:Label>&nbsp;<asp:DropDownList ID="ddlAccountHandler" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlAccountHandler_SelectedIndexChanged" Visible="False">
                        </asp:DropDownList>
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                        <asp:CheckBox ID="cbRetrospective" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="include events in last" AutoPostBack="True" 
                            oncheckedchanged="cbRetrospective_CheckedChanged" />
                        &nbsp;<asp:DropDownList ID="ddlCMRetrospectiveDays" runat="server" 
                            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" 
                            onselectedindexchanged="ddlCMRetrospectiveDays_SelectedIndexChanged" 
                            Enabled="False">
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem Selected="True">30</asp:ListItem>
                            <asp:ListItem>60</asp:ListItem>
                            <asp:ListItem>90</asp:ListItem>
                        </asp:DropDownList>
                          &nbsp;<asp:Label ID="lblDays" runat="server" Text="days" Font-Names="Verdana" 
                            Font-Size="XX-Small"/>
                        <asp:CheckBox ID="cbIncludeCancelledEvents" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="cancelled events" AutoPostBack="True" oncheckedchanged="cbIncludeCancelledEvents_CheckedChanged" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 14px">
                    </td>
                    <td colspan="2" style="height: 14px">
                        <asp:Label ID="lblLegendEvents" runat="server" Text="Events:" Visible="False" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                    <td style="height: 14px">
                    </td>
                    <td style="height: 14px">
                    </td>
                    <td style="height: 14px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:GridView ID="gvEvents" runat="server" CellPadding="2" Width="100%" Visible="False" OnRowDataBound="gvEvents_RowDataBound" AllowPaging="True" OnPageIndexChanging="gvEvents_PageIndexChanging" Font-Names="Verdana" Font-Size="XX-Small">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnReviewEvent" CommandArgument='<%# Container.DataItem("EventId")%>' OnClick="lnkbtnReviewEvent_Click" runat="server">review</asp:LinkButton>
                                        <asp:LinkButton ID="lnkbtnDeleteEvent" CommandArgument='<%# Container.DataItem("EventId")%>' OnClick="lnkbtnDeleteEvent_Click" OnClientClick='return confirm("Are you sure you want to delete this event?");' runat="server">delete</asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <EmptyDataTemplate>
                                no events found
                            </EmptyDataTemplate>
                            <PagerStyle Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td colspan="4">
                        <asp:Label ID="lblItemsPerPage" runat="server" 
                            Text="Items per page:" Font-Names="Verdana" 
                            Font-Size="XX-Small" Visible="False"></asp:Label>
                        <asp:DropDownList ID="ddlCMItemsPerPage" runat="server" AutoPostBack="True" 
                            Font-Names="Verdana" Font-Size="XX-Small" 
                            onselectedindexchanged="ddlCMItemsPerPage_SelectedIndexChanged" 
                            Visible="False">
                            <asp:ListItem Selected="True">10</asp:ListItem>
                            <asp:ListItem>25</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        <asp:Panel ID="pnlEvent" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 16%" align="right">
                    </td>
                    <td style="width: 33%">
                    </td>
                    <td style="width: 16%">
                    </td>
                    <td style="width: 33%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        &nbsp;<asp:Label ID="lblLegendEvent" runat="server" Text="Event Details:" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"/></td>
                    <td>
                    </td>
                    <td style="width: 450px">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer:"/></td>
                    <td>
                        <asp:Label ID="lblCustomer" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"/>
                        <asp:HiddenField ID="hidCustomerKey" runat="server" />
                    </td>
                    <td colspan="2">
                        <asp:Label ID="Label16" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Booked by"/>
                        <asp:Label ID="lblBookedBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/><asp:Label ID="Label17" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="&nbsp;on"/>
                        <asp:Label ID="lblBookedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Name:"/></td>
                    <td>
                        <asp:Label ID="lblEventName" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"/></td>
                    <td colspan="2">
                        <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Delivery date:"/>
                        <asp:Label ID="lblDeliveryDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"/>
                        &nbsp;
                        <asp:Label ID="Label26" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection date:"/>
                        <asp:Label ID="lblCollectionDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvContactName" ControlToValidate="tbContactName" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact Name:" ForeColor="Red"/>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="tbContactName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvContactPhone" ControlToValidate="tbContactPhone" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact Phone:" ForeColor="Red"/></td>
                    <td>
                        <asp:TextBox ID="tbContactPhone" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/></td>
                    <td align="right">
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact Mobile:"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbContactMobile" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactName2" runat="server" ControlToValidate="tbCMContactName2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" Visible="False"/>
                        <asp:Label ID="lblLegendCMContactName2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact name 2:" Visible="False"/>
                    </td>
                    <td>
                        <asp:LinkButton ID="lnkbtnCMAddSecondContact" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnCMAddSecondContact_Click">add 2nd contact</asp:LinkButton>                    
                        <asp:TextBox ID="tbCMContactName2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" Visible="False" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactMobile2" runat="server" ControlToValidate="tbCMContactMobile2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="lblLegendCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Contact mobile 2:"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr ID="trCMContactPhone2" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactPhone2" runat="server" ControlToValidate="tbCMContactPhone2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label15axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact phone 2:"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCMContactPhone2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        <asp:LinkButton ID="lnkbtnCMRemoveSecondContact" runat="server" 
                            OnClientClick='return confirm("Are you sure you want to remove the 2nd contact?");' 
                            onclick="lnkbtnCMRemoveSecondContact_Click" Font-Names="Verdana" 
                            Font-Size="XX-Small">remove 2nd contact</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvEventAddress1" ControlToValidate="tbEventAddress1" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Address 1:" ForeColor="Red"/>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="tbEventAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td >
                    </td>
                    <td align="right">
                        <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Address 2:"/></td>
                    <td>
                        <asp:TextBox ID="tbEventAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/></td>
                    <td align="right">
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event Address 3:"/></td>
                    <td>
                        <asp:TextBox ID="tbEventAddress3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvTown" ControlToValidate="tbTown" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Town:" ForeColor="Red"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvPostCode" ControlToValidate="tbPostcode" runat="server" ErrorMessage="#" Font-Size="XX-Small" ValidationGroup="CalendarManaged"/>                        
                        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Post Code:" ForeColor="Red"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="150px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                        &nbsp;<asp:LinkButton ID="lnkbtnCMAddressOutsideUK" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnCMAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trCMCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCountry" runat="server" 
                            ControlToValidate="ddlCMCountry" ErrorMessage="#" Font-Names="Verdana" 
                            Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa0" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCountry" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px">
                        </asp:DropDownList>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Delivery Time:" ForeColor="Red"/></td>
                    <td>
                        <asp:DropDownList ID="ddlDeliveryTime" runat="server" Font-Names="Verdana" Font-Size="XX-Small" BackColor="LightGoldenrodYellow">
                            <asp:ListItem>9.00am</asp:ListItem>
                            <asp:ListItem>10.30am</asp:ListItem>
                            <asp:ListItem>12.00 noon</asp:ListItem>
                            <asp:ListItem>Other times pls specify in Special Instructions</asp:ListItem>
                        </asp:DropDownList></td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvPreciseDeliveryPoint" ControlToValidate="tbPreciseDeliveryPoint" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Exact Delivery Point:" ForeColor="Red"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPreciseDeliveryPoint" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" MaxLength="100" BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection Time:" ForeColor="Red"/></td>
                    <td>
                        <asp:DropDownList ID="ddlCollectionTime" runat="server" Font-Names="Verdana" Font-Size="XX-Small" BackColor="LightGoldenrodYellow">
                            <asp:ListItem>9.00am - 10.00am</asp:ListItem>
                            <asp:ListItem>10.00am - 11.00am</asp:ListItem>
                            <asp:ListItem>11.00am - 12.00 noon</asp:ListItem>
                            <asp:ListItem>12.00 noon - 1.00pm</asp:ListItem>
                            <asp:ListItem>1.00pm - 2.00pm</asp:ListItem>
                            <asp:ListItem>2.00pm - 3.00pm</asp:ListItem>
                            <asp:ListItem>3.00pm - 4.00pm</asp:ListItem>
                            <asp:ListItem>4.00pm - 5.00pm</asp:ListItem>
                            <asp:ListItem>5.00pm - 6.00pm</asp:ListItem>
                            <asp:ListItem>Other - contact Transworld</asp:ListItem>
                        </asp:DropDownList></td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvPreciseCollectionPoint" ControlToValidate="tbPreciseCollectionPoint" runat="server" Font-Size="XX-Small" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Exact Collection Point:" ForeColor="Red"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPreciseCollectionPoint" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" MaxLength="100" BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbDifferentCollectionAddress" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="collect from a different address" AutoPostBack="True" OnCheckedChanged="cbDifferentCollectionAddress_CheckedChanged" />
                    </td>
                    <td align="right">
                    </td>
                    <td style="width: 450px">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trCollectionAddress1" runat="server" visible="false">
                    <td style="height: 18px">
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCollectionAddress1" 
                            ControlToValidate="tbCollectionAddress1" runat="server" Font-Size="XX-Small" 
                            ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label78" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection address 1:" ForeColor="Red"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCollectionAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label79" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection address 2:"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCollectionAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trCollectionAddress2" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCollectionTown" runat="server" 
                            ControlToValidate="tbCollectionTown" ErrorMessage="#" Font-Size="XX-Small" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label80" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection town:" ForeColor="Red"/></td>
                    <td>
                        <asp:TextBox ID="tbCollectionTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCollectionPostcode" runat="server" 
                            ControlToValidate="tbCollectionPostcode" ErrorMessage="#" Font-Size="XX-Small" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label81" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection post code:" ForeColor="Red"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCollectionPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="150px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                        &nbsp;<asp:LinkButton ID="lnkbtnCMCollectionAddressOutsideUK" runat="server" 
                            Font-Size="XX-Small" Font-Names="Verdana" 
                            onclick="lnkbtnCMCollectionAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr ID="trCMCollectionCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCollectionCountry" runat="server" ControlToValidate="ddlCMCollectionCountry" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCollectionCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px"/>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label82" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer reference:"/></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbCustomerReference" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Special Instructions:"/>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="tbSpecialInstructions" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="99%" MaxLength="180" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendProduct" runat="server" Text="Product:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td colspan="3">
                        <asp:GridView ID="gvItems" runat="server" CellPadding="2" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3">
                        <asp:Button ID="btnSaveChanges" runat="server" Text="save event changes" OnClick="btnSaveChanges_Click" />
                        <asp:Button ID="btnPickEventProducts" runat="server" OnClick="btnPickEventProducts_Click"
                            Text="pick products for event" />&nbsp;
                        <asp:Label ID="lblProductPickedFlag" runat="server" Font-Bold="True" 
                            Font-Names="Verdana" Font-Size="X-Small" ForeColor="#006600" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                      <hr />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td valign="bottom">
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="New note:" Font-Bold="True"/></td>
                    <td colspan="2" valign="middle">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 27px">
                    </td>
                    <td align="right" style="height: 27px">
                    </td>
                    <td valign="bottom" colspan="3" style="height: 27px">
                        <asp:TextBox ID="tbNote" runat="server" Width="70%" MaxLength="180" Font-Names="Verdana" Font-Size="XX-Small" BackColor="LightGoldenrodYellow"></asp:TextBox>
                        <asp:CheckBox ID="cbCustomerVisible" runat="server" Checked="True" Text="customer-visible" BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;
                        <asp:Button ID="btnAddNote" runat="server" Text="add note" OnClick="btnAddNote_Click" /></td>
                    <td style="height: 27px">
                    </td>
                </tr>
                <tr id="trNotes" runat="server">
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3"><asp:Label ID="Label25" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Notes:" Font-Bold="True"/>
                        <asp:GridView ID="gvNotes" runat="server" CellPadding="2" Width="100%" AllowPaging="True" PageSize="6" OnPageIndexChanging="gvNotes_PageIndexChanging" Font-Names="Verdana" Font-Size="XX-Small">
                            <EmptyDataTemplate>
                                no notes
                            </EmptyDataTemplate>
                            <PagerStyle Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                        </asp:GridView>
                        &nbsp;
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 21px">
                    </td>
                    <td align="right" style="height: 21px">
                    </td>
                    <td colspan="3" style="height: 21px">
                        <asp:LinkButton ID="lnkbtnShowHideNotes" runat="server" OnClick="lnkbtnShowHideNotes_Click" Font-Names="Verdana" Font-Size="XX-Small">hide notes</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtnRefreshNotes" runat="server" OnClick="lnkbtnRefreshNotes_Click" Font-Names="Verdana" Font-Size="XX-Small">refresh notes</asp:LinkButton></td>
                    <td style="height: 21px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                      <hr />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3"><asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="New reminder:" Font-Bold="True"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3" style="font-size: xx-small; font-family: Verdana">
                        &nbsp;<asp:Label ID="Label21" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Send on "/>
                        <asp:TextBox ID="tbEmailReminderDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="75px" BackColor="LightGoldenrodYellow"/> (eg 30-Jun-2009)
                            <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbEmailReminderDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="imgEmailDateCalendar" alt="" style="border: none"
                                 src="~/images/SmallCalendar.gif"
                                 runat="server"
                              IE:visible="true"
                                 visible="false"
                               /></a>
                        at <asp:DropDownList ID="ddlEmailReminderTime" runat="server" Font-Names="Verdana" Font-Size="XX-Small" BackColor="LightGoldenrodYellow">
                            <asp:ListItem Value="0800" Text="8.00am"/>
                            <asp:ListItem Value="0815" Text="8.15am"/>
                            <asp:ListItem Value="0830" Text="8.30am"/>
                            <asp:ListItem Value="0845" Text="8.45am"/>
                            <asp:ListItem Value="0900" Text="9.00am"/>
                            <asp:ListItem Value="0915" Text="9.15am"/>
                            <asp:ListItem Value="0930" Text="9.30am"/>
                            <asp:ListItem Value="0945" Text="9.45am"/>
                            <asp:ListItem Value="1000" Text="10.00am"/>
                            <asp:ListItem Value="1015" Text="10.15am"/>
                            <asp:ListItem Value="1030" Text="10.30am"/>
                            <asp:ListItem Value="1045" Text="10.45am"/>
                            <asp:ListItem Value="1100" Text="11.00am"/>
                            <asp:ListItem Value="1115" Text="11.15am"/>
                            <asp:ListItem Value="1130" Text="11.30am"/>
                            <asp:ListItem Value="1145" Text="11.45am"/>
                            <asp:ListItem Value="1200" Text="12.00 noon"/>
                            <asp:ListItem Value="1215" Text="12.15pm"/>
                            <asp:ListItem Value="1230" Text="12.30pm"/>
                            <asp:ListItem Value="1245" Text="12.45pm"/>
                            <asp:ListItem Value="1300" Text="1.00pm"/>
                            <asp:ListItem Value="1315" Text="1.15pm"/>
                            <asp:ListItem Value="1330" Text="1.30pm"/>
                            <asp:ListItem Value="1345" Text="1.45pm"/>
                            <asp:ListItem Value="1400" Text="2.00pm"/>
                            <asp:ListItem Value="1415" Text="2.15pm"/>
                            <asp:ListItem Value="1430" Text="2.30pm"/>
                            <asp:ListItem Value="1445" Text="2.45pm"/>
                            <asp:ListItem Value="1500" Text="3.00pm"/>
                            <asp:ListItem Value="1515" Text="3.15pm"/>
                            <asp:ListItem Value="1530" Text="3.30pm"/>
                            <asp:ListItem Value="1545" Text="3.45pm"/>
                            <asp:ListItem Value="1600" Text="4.00pm"/>
                            <asp:ListItem Value="1615" Text="4.15pm"/>
                            <asp:ListItem Value="1630" Text="4.30pm"/>
                            <asp:ListItem Value="1645" Text="4.45pm"/>
                            <asp:ListItem Value="1700" Text="5.00pm"/>
                            <asp:ListItem Value="1715" Text="5.15pm"/>
                            <asp:ListItem Value="1730" Text="5.30pm"/>
                            <asp:ListItem Value="1745" Text="5.45pm"/>
                            <asp:ListItem Value="1800" Text="6.00pm"/>
                            <asp:ListItem Value="1815" Text="6.15pm"/>
                            <asp:ListItem Value="1830" Text="6.30pm"/>
                            <asp:ListItem Value="1845" Text="6.45pm"/>
                        </asp:DropDownList>&nbsp;to
                        <asp:TextBox ID="tbEmailReminderAddr" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" BackColor="LightGoldenrodYellow"/>&nbsp;
                        <asp:LinkButton ID="lnkbtnEmailAddrMe" runat="server" 
                            onclick="lnkbtnEmailAddrMe_Click">me</asp:LinkButton>
                        </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td valign="bottom" colspan="3">
                        &nbsp;<asp:Label ID="Label22" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Text:"/>
                        <asp:TextBox ID="tbEmailReminderText" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="75%" BackColor="LightGoldenrodYellow"/>&nbsp;
                        <asp:Button ID="btnAddReminder" runat="server" Text="add reminder" OnClick="btnAddReminder_Click" /></td>
                    <td>
                    </td>
                </tr>
                <tr id="trReminders" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        </td>
                    <td colspan="3"><asp:Label ID="Label24" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Reminders:" Font-Bold="True"/><asp:GridView ID="gvReminders" runat="server" CellPadding="2" Width="100%" AllowPaging="True" PageSize="6" OnRowDataBound="gvReminders_RowDataBound" OnPageIndexChanging="gvReminders_PageIndexChanging" Font-Names="Verdana" Font-Size="XX-Small">
                            <EmptyDataTemplate>
                                <span style="color: red">NO email reminders set</span>
                            </EmptyDataTemplate>
                            <PagerStyle Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnDeleteReminder" CommandArgument='<%# Container.DataItem("key")%>' OnClientClick='return confirm("Are you sure you want to delete this reminder?");' runat="server" Text="delete reminder" OnClick="lnkbtnDeleteReminder_Click" Font-Names="Verdana" Font-Size="XX-Small" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                        &nbsp;
                    </td>
                    <td>
                    </td>
                </tr>
                <tr runat="server">
                    <td style="height: 21px">
                    </td>
                    <td align="right" style="height: 21px">
                    </td>
                    <td colspan="3" style="height: 21px">
                        <asp:LinkButton ID="lnkbtnShowHideReminders" runat="server" OnClick="lnkbtnShowHideReminders_Click" Font-Names="Verdana" Font-Size="XX-Small">hide reminders</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtrnRefreshReminders" runat="server" OnClick="lnkbtrnRefreshReminders_Click" Font-Names="Verdana" Font-Size="XX-Small">refresh reminders</asp:LinkButton></td>
                    <td style="height: 21px">
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <hr />
        <asp:Panel ID="pnlSiteFeatures" runat="server" Width="100%" Visible="True">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label19" runat="server" Text="Site administrator:" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label></td>
                    <td>
                    </td>
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
                        <asp:DropDownList ID="ddlSiteAdministrator" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlSiteAdministrator_SelectedIndexChanged" AutoPostBack="True"/>
                        <asp:HiddenField ID="hidSiteAdministratorKey" runat="server" />
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                        <asp:Label ID="Label27" runat="server" Text="Site Features" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label>
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%" align="right">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label35" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Postcode lookup:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbPostcodeLookup" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td align="right">
                        <asp:Label ID="Label36" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Calendar management:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbCalendarManagement" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label44" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Product owners:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbProductOwners" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                    </td>
                    <td align="right">
                        <asp:Label ID="Label45" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Category count:"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCategoryCount" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="50px"/>
                        <asp:RequiredFieldValidator ID="rfvCategoryCount" runat="server" ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small" ControlToValidate="tbCategoryCount"/>
                        <asp:RangeValidator ID="rvCategoryCount" runat="server" ErrorMessage="2 or 3" Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="3" MinimumValue="2" Type="Integer" ControlToValidate="tbCategoryCount"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right">
                        <asp:Label ID="Label46" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Web form:"/>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbWebForm" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbWebForm_CheckedChanged" AutoPostBack="True" />
                    </td>
                    <td align="right">
                        <asp:Label ID="Label57" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Session timeout (mins):"></asp:Label>
                    </td>
                    <td>
                        <asp:TextBox ID="tbSessionTimeout" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="50px"/>
                        <asp:RequiredFieldValidator ID="rfvSessionTimeout" runat="server" ControlToValidate="tbSessionTimeout" ErrorMessage="<< required" Font-Names="Verdana" Font-Size="XX-Small"/>
                        <asp:RangeValidator ID="rvSessionTimeout" runat="server" ControlToValidate="tbSessionTimeout" ErrorMessage="1 - 999" Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="999" MinimumValue="1" Type="Integer"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label47" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="On-demand products:"></asp:Label></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbOnDemandProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOnDemandProducts_CheckedChanged" />
                        <asp:CheckBox ID="cbPrintOndemandTab" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" /></td>
                    <td align="right">
                        <asp:Label ID="Label75" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Custom letters:"></asp:Label></td>
                    <td><asp:CheckBox ID="cbCustomLetters" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label76" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Advanced user permissions:"></asp:Label></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbUserPermissions" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOnDemandProducts_CheckedChanged" /></td>
                    <td align="right">
                        <asp:Label ID="Label77" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="Secure file upload:"></asp:Label></td>
                    <td><asp:CheckBox ID="cbFileUpload" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbFileUpload_CheckedChanged" />
                        &nbsp;&nbsp;
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td style="height: 22px">
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendFileUploadNotificationEmailAddresses" runat="server" Text="Email file upload notifications to (separate with commas):" Font-Names="Verdana" Font-Size="XX-Small" Visible="False"/></td>
                    <td>
                        <asp:TextBox ID="tbFileUploadNotificationEmailAddresses" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" Visible="False"/></td>
                    <td>
                    </td>
                </tr>
                <tr id="trAttemAccess" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label69" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="ATTEM access for customer:"></asp:Label></td>
                    <td colspan="3">
                        <asp:Label ID="lblCustomerName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" ForeColor="Red"></asp:Label>&nbsp;
                        <asp:Label ID="Label70" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="ATTEM user name:"></asp:Label>
                        <asp:TextBox ID="tbAttemUserName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="75px"/>
                        <asp:Label ID="Label71" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="ATTEM password:"></asp:Label><asp:TextBox ID="tbAttemPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="75px"/>
                        <asp:Label ID="Label72" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="ATTEM customer name:"></asp:Label><asp:TextBox ID="tbAttemCustomerName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="75px"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="Label48" runat="server" Text="STOCK BOOKING OPTIONS" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label50" runat="server" Text="Cust Ref 1 options:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbCustRef1IsVisible" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is visible?" AutoPostBack="True" OnCheckedChanged="cbCustRef1IsVisible_CheckedChanged" />
                        <asp:CheckBox ID="cbCustRef1IsMandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is mandatory?" /></td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label49" runat="server" Text="Cust Ref 1 label:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbCustRef1Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px" MaxLength="20"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label51" runat="server" Text="Cust Ref 2 options:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbCustRef2IsVisible" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is visible?" AutoPostBack="True" OnCheckedChanged="cbCustRef2IsVisible_CheckedChanged" />
                        <asp:CheckBox ID="cbCustRef2IsMandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is mandatory?" /></td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label54" runat="server" Text="Cust Ref 2 label:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbCustRef2Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px" MaxLength="20"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label52" runat="server" Text="Cust Ref 3 options:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbCustRef3IsVisible" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is visible?" AutoPostBack="True" OnCheckedChanged="cbCustRef3IsVisible_CheckedChanged" />
                        <asp:CheckBox ID="cbCustRef3IsMandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is mandatory?" /></td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label55" runat="server" Text="Cust Ref 3 label:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbCustRef3Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px" MaxLength="20"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label53" runat="server" Text="Cust Ref 4 options:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbCustRef4IsVisible" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is visible?" AutoPostBack="True" OnCheckedChanged="cbCustRef4IsVisible_CheckedChanged" />
                        <asp:CheckBox ID="cbCustRef4IsMandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Is mandatory?" /></td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label56" runat="server" Text="Cust Ref 4 label:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbCustRef4Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px" MaxLength="20"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td style="height: 22px">
                        <asp:Label ID="Label40" runat="server" Text="COURIER BOOKING OPTIONS" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                    </td>
                    <td style="height: 22px">
                    </td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label37" runat="server" Text="Use label printer:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbUseLabelPrinter" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td style="height: 22px" align="right">
                        <asp:Label ID="Label39" runat="server" Text="Default consignment description:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbDefaultDescription" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label38" runat="server" Text="Search company name only:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbSearchCompanyNameOnly" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
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
                    <td align="right">
                        <asp:Label ID="Label29" runat="server" Text="Make Ref1 mandatory:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbMakeRef1Mandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td align="right">
                    <asp:Label ID="Label30" runat="server" Text="Make Ref2 mandatory:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    <asp:CheckBox ID="cbMakeRef2Mandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label28" runat="server" Text="Ref1 label:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td style="height: 22px">
                        <asp:TextBox ID="tbRef1Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/>
                    </td>
                    <td align="right" style="height: 22px">
                    <asp:Label ID="Label33" runat="server" Text="Ref2 label:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                    <asp:TextBox ID="tbRef2Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:Label ID="Label31" runat="server" Text="Make Ref3 mandatory:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td><asp:CheckBox ID="cbMakeRef3Mandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td align="right">
                    <asp:Label ID="Label32" runat="server" Text="Make Ref4 mandatory:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    <asp:CheckBox ID="cbMakeRef4Mandatory" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:Label ID="Label34" runat="server" Text="Ref3 label:" Font-Names="Verdana" Font-Size="XX-Small"/>&nbsp;</td>
                    <td>
                        <asp:TextBox ID="tbRef3Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td align="right">
                    <asp:Label ID="Label41" runat="server" Text="Ref4 label:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    <asp:TextBox ID="tbRef4Label" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label43" runat="server" Text="Third party collection key:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbThirdPartyCollectionKey" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="50px"/>
                        <asp:RequiredFieldValidator ID="rfvThirdPartyCollectionKey" runat="server" ControlToValidate="tbThirdPartyCollectionKey"
                            ErrorMessage="< required" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator>
                        <asp:RangeValidator ID="rvThirdPartyCollectionKey" runat="server" ControlToValidate="tbThirdPartyCollectionKey"
                            ErrorMessage="!!!" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="99999" MinimumValue="-1" Type="Integer"></asp:RangeValidator></td>
                    <td align="right">
                        <asp:Label ID="Label42" runat="server" Text="Hide collection button:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbHideCollectionButton" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm01" runat="server">
                    <td>
                    </td>
                    <td>
                        <asp:Label ID="Label58" runat="server" Text="WEB FORM OPTIONS" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm02" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label59" runat="server" Text="Customer:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:DropDownList ID="ddlWebFormCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlWebFormCustomer_SelectedIndexChanged" AutoPostBack="True">
                        </asp:DropDownList></td>
                    <td align="right">
                        <asp:Label ID="Label60" runat="server" Text="Generic user:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:DropDownList ID="ddlWebFormGenericUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        </asp:DropDownList></td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm03" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label61" runat="server" Text="Web form page title:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:TextBox ID="tbWebFormPageTitle" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td align="right">
                        <asp:Label ID="Label62" runat="server" Text="Web form logo:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:TextBox ID="tbWebFormLogoImage" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm04" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label63" runat="server" Text="Web form upper heading:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:TextBox ID="tbWebFormTopLegend" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td align="right">
                        <asp:Label ID="Label64" runat="server" Text="Web form lower heading:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:TextBox ID="tbWebFormBottomLegend" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="250px"/></td>
                    <td>
                    </td>
                </tr>
                <tr runat="server">
                    <td style="height: 22px">
                    </td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label73" runat="server" Text="Show products with zero qty:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbWebFormShowZeroQuantity" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td align="right" style="height: 22px">
                        <asp:Label ID="Label74" runat="server" Text="Collect user details for zero qty products:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td style="height: 22px">
                        <asp:CheckBox ID="cbWebFormZeroStockNotification" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td style="height: 22px">
                    </td>
                </tr>
                <tr id="trWebForm05" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label68" runat="server" Text="Show price:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                        <asp:CheckBox ID="cbWebFormShowPrice" runat="server" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm06" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label65" runat="server" Text="Web form home page:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td colspan="3">
                        <FCKeditorV2:FCKeditor ID="FCKedWebFormHomePage" runat="server" ToolbarSet="CourierSoftware" BasePath="./fckeditor/" Height="100px" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm07" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label67" runat="server" Text="Web form address page:" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td colspan="3">
                        <FCKeditorV2:FCKeditor ID="FCKedWebFormAddressPage" runat="server" ToolbarSet="CourierSoftware" BasePath="./fckeditor/" Height="100px" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trWebForm08" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label66" runat="server" Text="Web form help page:" Font-Names="Verdana" Font-Size="XX-Small"/>
                    </td>
                    <td colspan="3">
                        <FCKeditorV2:FCKeditor ID="FCKedWebFormHelpPage" runat="server" ToolbarSet="CourierSoftware" BasePath="./fckeditor/" Height="100px" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:Button ID="btnSaveSiteFeatureChanges" runat="server" Text="save site feature changes" OnClick="btnSaveSiteFeatureChanges_Click" /></td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
     </form>
</body>
</html>
