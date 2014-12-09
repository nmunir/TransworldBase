<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data " %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Drawing.Color" %>
<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Const CURRENCY_STERLING As Integer = 123

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Call GetSiteFeatures()
            If pbHideCollectionButton Then
                btnBookACollection.Visible = False
            End If
    
            Dim iThisMinute As Integer = Minute(Now)
            Dim iThisHour As Integer = Hour(Now)
            Dim iThisDay As Integer = Day(Now)
            Dim iThisMonth As Integer = DatePart(DateInterval.Month, Now)
            Dim iThisYear As Integer = Year(Now)
    
            Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)
    
            If iThisMinute >= 0 And iThisMinute < 15 Then
                drop_MinuteReady.SelectedIndex = 1
            ElseIf iThisMinute >= 15 And iThisMinute < 30 Then
                drop_MinuteReady.SelectedIndex = 2
            ElseIf iThisMinute >= 30 And iThisMinute < 45 Then
                drop_MinuteReady.SelectedIndex = 3
            ElseIf iThisMinute >= 45 And iThisMinute <= 59 Then
                iThisHour = iThisHour + 1
                drop_MinuteReady.SelectedIndex = 0
            End If
    
            drop_YearReady.Items.Add(iThisYear)
            drop_YearReady.Items.Add(iThisYear + 1)
    
            drop_HourReady.SelectedIndex = iThisHour - 1
            drop_DayReady.SelectedIndex = iThisDay - 1
            drop_MonthReady.SelectedIndex = iThisMonth - 1
            drop_YearReady.SelectedValue = iThisYear
    
            Call GetUsersCompanyDetails()
            Call PopulateCountryDropDowns()
            If plThirdPartyCollectionKey > 0 Then
                pnlCollectionStatus.Visible = False
            Else
                pnlCollectionStatus.Visible = True
                BindTodaysCollectionsGrid()
            End If
            Call BindConsAwaitingCollGrid()
            Call ShowCourierBookingStatus()
            Call ResetConsignee()

            txtSearchCriteriaAddress.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
            txtDescription.Attributes.Add("onBlur", "chkDocs();")
        End If
        Response.Buffer = True
        Call SetTitle()
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Courier Booking"
    End Sub
    
    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub GetSiteFeatures()
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
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pbUseLabelPrinter = dr("UseLabelPrinter")
        pbSearchCompanyNameOnly = dr("SearchCompanyNameOnly")
        psDefaultDescription = dr("DefaultDescription")
        pbMakeRef1Mandatory = dr("MakeRef1Mandatory")
        lblRef1.Text = dr("Ref1Label")
        pbMakeRef2Mandatory = dr("MakeRef2Mandatory")
        lblRef2.Text = dr("Ref2Label")
        pbMakeRef3Mandatory = dr("MakeRef3Mandatory")
        lblRef3.Text = dr("Ref3Label")
        pbMakeRef4Mandatory = dr("MakeRef4Mandatory")
        lblRef4.Text = dr("Ref4Label")
        plThirdPartyCollectionKey = dr("ThirdPartyCollectionKey")
        pbHideCollectionButton = dr("HideCollectionButton")
    End Sub

    Protected Sub HideAllPanels()
        pnlCourierBookingStatus.Visible = False
        pnlSearchAddressBook.Visible = False
        pnlCreateConsignment.Visible = False
        pnlEditConsignment.Visible = False
        pnlRequestSupplies.Visible = False
        pnlConfirmDeleteConsignment.Visible = False
        pnlConfirmNewConsignment.Visible = False
        pnlConfirmDeleteCollection.Visible = False
        pnlCreateAdHocCollection.Visible = False
        pnlConfirmNewCourierCollection.Visible = False
        pnlAmendCollection.Visible = False
    End Sub
    
    Protected Sub ShowCourierBookingStatus()
        Call HideAllPanels()
        Call ClearPreAlertAddresses()
        pnlCourierBookingStatus.Visible = True
    End Sub
    
    Protected Sub ShowAddressBook()
        Call HideAllPanels()
        pnlSearchAddressBook.Visible = True
        txtSearchCriteriaAddress.Text = String.Empty
        lblAddressMessage.Text = String.Empty
        txtSearchCriteriaAddress.Focus ()
    End Sub
    
    Protected Sub ShowCreateConsignment()
        Call HideAllPanels()
        pnlCreateConsignment.Visible = True
    End Sub
    
    Protected Sub ShowEditConsignment()
        Call HideAllPanels()
        pnlEditConsignment.Visible = True
    End Sub
    
    Protected Sub ShowRequestSupplies()
        Call HideAllPanels()
        pnlRequestSupplies.Visible = True
    End Sub
    
    Protected Sub ShowConfirmDeleteConsignment()
        Call HideAllPanels()
        pnlConfirmDeleteConsignment.Visible = True
    End Sub
    
    Protected Sub ShowConfirmDeleteCollection()
        Call HideAllPanels()
        pnlConfirmDeleteCollection.Visible = True
    End Sub
    
    Protected Sub ShowConfirmNewConsignment()
        Call HideAllPanels()
        pnlConfirmNewConsignment.Visible = True
    End Sub
    
    Protected Sub ShowCreateAdHocCollection()
        Call HideAllPanels()
        pnlCreateAdHocCollection.Visible = True
    End Sub
    
    Protected Sub ShowConfirmNewCourierCollection()
        Call HideAllPanels()
        pnlConfirmNewCourierCollection.Visible = True
    End Sub
    
    Protected Sub ShowAmendCollection()
        Call HideAllPanels()
        pnlAmendCollection.Visible = True
    End Sub
    
    Protected Sub DisplayCourierBookingStatus()
        BindConsAwaitingCollGrid()
        ShowCourierBookingStatus()
    End Sub
    
    Protected Sub CancelConsignment()
        ResetConsignmentForm()
        BindConsAwaitingCollGrid()
        ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btn_GetFromAddressBook_click(ByVal s As Object, ByVal e As EventArgs)
        ShowAddressBook()
    End Sub
    
    Protected Sub GoBackToConsignment()
        pbIsCreatingNewConsignment = True
        dgConsignmentsAwaitingCollection.CurrentPageIndex = 0
        ShowCreateConsignment()
    End Sub
    
    Protected Sub CreateConsignment()
        If pbMakeRef1Mandatory = True Then
            lblRef1.ForeColor = Drawing.Color.Red
            validator_Ref1.Enabled = True
        Else
            lblRef1.ForeColor = Drawing.Color.Gray
            validator_Ref1.Enabled = False
        End If
        If pbMakeRef2Mandatory = True Then
            lblRef2.ForeColor = Drawing.Color.Red
            validator_Ref2.Enabled = True
        Else
            lblRef2.ForeColor = Drawing.Color.Gray
            validator_Ref2.Enabled = False
        End If
        If pbMakeRef3Mandatory = True Then
            lblRef3.ForeColor = Drawing.Color.Red
            validator_Ref3.Enabled = True
        Else
            lblRef3.ForeColor = Drawing.Color.Gray
            validator_Ref3.Enabled = False
        End If
        If pbMakeRef4Mandatory = True Then
            lblRef4.ForeColor = Drawing.Color.Red
            validator_Ref4.Enabled = True
        Else
            lblRef4.ForeColor = Drawing.Color.Gray
            validator_Ref4.Enabled = False
        End If
        txtDescription.Text = psDefaultDescription
        pbIsCreatingNewConsignment = True
        dgConsignmentsAwaitingCollection.CurrentPageIndex = 0
        Call ShowCreateConsignment()
    End Sub
    
    Protected Sub btnConsignmentConfirmationYes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim lChosenCollectionKey As Long
        Dim iCount As Integer
        If psBookingState = "ZERO" Then
            pbIsCreatingBookingForConsignment = True
            Call ShowCreateAdHocCollection()
        ElseIf psBookingState = "ONE" Then
            For Each dgi In grid_1AvailableCollection.Items
                Dim cellCollectionKey As TableCell = dgi.Cells(0)
                lChosenCollectionKey = CLng( cellCollectionKey.Text)
                Call AssociateConsignmentWithCollection(plConsignmentKey, lChosenCollectionKey)
            Next dgi
            If plThirdPartyCollectionKey > 0 Then
                pnlCollectionStatus.Visible = False
            Else
                pnlCollectionStatus.Visible = True
                Call BindTodaysCollectionsGrid()
            End If
            Call BindConsAwaitingCollGrid()
            Call ShowCourierBookingStatus()
        ElseIf psBookingState = "MANY" Then
            lblGridSelection.Text = ""
            For Each dgi In grid_AvailableCollections.Items
                cb = CType(dgi.Cells(7).Controls(1), CheckBox)
                If cb.Checked = True Then
                    iCount = iCount + 1
                End If
            Next dgi
            If iCount = 1 Then
                For Each dgi In grid_AvailableCollections.Items
                    cb = CType(dgi.Cells(7).Controls(1), CheckBox)
                    If cb.Checked = True Then
                        Dim cellCollectionKey As TableCell = dgi.Cells(0)
                        lChosenCollectionKey = CLng(cellCollectionKey.Text)
                        AssociateConsignmentWithCollection(plConsignmentKey, lChosenCollectionKey)
                    End If
                Next dgi
            ElseIf iCount > 1 Then
                lblGridSelection.Text = "Please ensure only one box is ticked."
                Exit Sub
            ElseIf iCount = 0 Then
                lblGridSelection.Text = "Please check one box to select your chosen collection."
                Exit Sub
            End If
            If plThirdPartyCollectionKey > 0 Then
                pnlCollectionStatus.Visible = False
            Else
                pnlCollectionStatus.Visible = True
                Call BindTodaysCollectionsGrid()
            End If
            Call BindConsAwaitingCollGrid()
            Call ShowCourierBookingStatus()
        End If
    End Sub
    
    Protected Sub SubmitNewConsignment()
        Page.Validate("vgCreateConsignment")
        If Page.IsValid Then
            Dim lKey As Long
            If pbIsCreatingNewConsignment Then
                lKey = AddNewConsignment()
                If lKey > 0 Then
                    If chk_SaveAddress.Checked = True Then
                        Call AddAddressToPersonalAddressBook()
                    End If
                    Call ResetConsignmentForm()
                    Call BindAvailableCollectionsGrid()
                    Call ShowConfirmNewConsignment()
                Else
                    Server.Transfer("error.aspx")
                End If
            Else
                Call ShowCourierBookingStatus()
            End If
        End If
    End Sub
    
    Protected Sub btnConsignmentConfirmationNo_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindConsAwaitingCollGrid()
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btnConfirmCancelCollection_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        dgTodaysCollections.CurrentPageIndex = 0
        Call CancelCollection(CLng(lblCollectionNumber.Text ))
        If plThirdPartyCollectionKey > 0 Then
            pnlCollectionStatus.Visible = False
        Else
            pnlCollectionStatus.Visible = True
            Call BindTodaysCollectionsGrid()
        End If
        Call BindConsAwaitingCollGrid()
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btnAmendCollectionSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs )
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim lChosenCollectionKey As Long
        Dim iCount As Integer
        If psBookingState = "ZERO" Then
            pbIsCreatingBookingForConsignment = True
            Call ShowCreateAdHocCollection()
        ElseIf psBookingState = "ONE" Then
            For Each dgi In grid_OneCollection.Items
                Dim cellCollectionKey As TableCell = dgi.Cells (0)
                lChosenCollectionKey = CLng(cellCollectionKey.Text)
                Call AssociateConsignmentWithCollection(plConsignmentKey, lChosenCollectionKey)
            Next dgi
            If plThirdPartyCollectionKey > 0 Then
                pnlCollectionStatus.Visible = False
            Else
                pnlCollectionStatus.Visible = True
                Call BindTodaysCollectionsGrid()
            End If
            Call BindConsAwaitingCollGrid()
            Call ShowCourierBookingStatus()
        ElseIf psBookingState = "MANY" Then
            lblAmendedGridSelection.Text = ""
            For Each dgi In grid_ChooseFromCollections.Items
                cb = CType(dgi.Cells(7).Controls(1), CheckBox)
                If cb.Checked = True Then
                    iCount = iCount + 1
                End If
            Next dgi
            If iCount = 1 Then
                For Each dgi In grid_ChooseFromCollections.Items
                    cb = CType(dgi.Cells(7).Controls(1), CheckBox)
                    If cb.Checked = True Then
                        Dim cellCollectionKey As TableCell = dgi.Cells(0)
                        lChosenCollectionKey = CLng(cellCollectionKey.Text)
                        Call AssociateConsignmentWithCollection(plConsignmentKey, lChosenCollectionKey)
                    End If
                Next dgi
            ElseIf iCount > 1 Then
                lblAmendedGridSelection.Text = "Please ensure only one box is ticked"
                Exit Sub
            ElseIf iCount = 0 Then
                lblAmendedGridSelection.Text = "Please check one box to select your chosen collection"
                Exit Sub
            End If
            If plThirdPartyCollectionKey > 0 Then
                pnlCollectionStatus.Visible = False
            Else
                pnlCollectionStatus.Visible = True
                Call BindTodaysCollectionsGrid()
            End If
            Call BindConsAwaitingCollGrid()
            Call ShowCourierBookingStatus()
        End If
    End Sub
    
    Protected Sub btnAmendCollectionCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btnUnBookConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        UnBookCollection(plConsignmentKey)
        If plThirdPartyCollectionKey > 0 Then
            pnlCollectionStatus.Visible = False
        Else
            pnlCollectionStatus.Visible = True
            Call BindTodaysCollectionsGrid()
        End If
        Call BindConsAwaitingCollGrid()
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub SubmitCourierCollection()
        Call AddCourierBooking()
        Call ShowConfirmNewCourierCollection()
    End Sub
    
    Protected Sub btn_DeletelConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DeleteConsignment(CLng(lblConsignmentNumberToDelete.Text))
        Call BindConsAwaitingCollGrid()
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btn_GotoSubmitNewConsignment_Click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        Call ShowConfirmNewConsignment()
    End Sub
    
    Protected Sub btn_CancelCreateConsignment_Click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        Call ResetConsignmentForm()
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btn_RefreshStatusPage_click(ByVal s As Object, ByVal e As EventArgs)
        If plThirdPartyCollectionKey > 0 Then
            pnlCollectionStatus.Visible = False
        Else
            pnlCollectionStatus.Visible = True
            Call BindTodaysCollectionsGrid()
        End If
        Call BindConsAwaitingCollGrid()
    End Sub
    
    Protected Sub ShowAllAddresses()
        txtSearchCriteriaAddress.Text = ""
        dgAddressBook.CurrentPageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub SearchAddresses()
        dgAddressBook.CurrentPageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub btn_SubmitSuppliesRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If chk_WindowEnvelopes.Checked = False And chk_LargeFlyerBags.Checked = False And chk_SmallFlyerBags.Checked = False Then
            WebMsgBox.Show("Please check one or more boxes to receive your supplies.")
        Else
            lblSuppliesMessage.Text = ""
            If chk_LargeFlyerBags.Checked = True Then
                Call AddRequestForSupplies("LARGE FLYER BAGS")
                chk_LargeFlyerBags.Checked = False
            End If
            If chk_SmallFlyerBags.Checked = True Then
                Call AddRequestForSupplies("SMALL FLYER BAGS")
                chk_SmallFlyerBags.Checked = False
            End If
            If chk_WindowEnvelopes.Checked = True Then
                Call AddRequestForSupplies("WINDOW ENVELOPES")
                chk_WindowEnvelopes.Checked = False
            End If
            WebMsgBox.Show("We have received your request for supplies, which we will process as soon as possible. Thank you.")
            Call ShowCourierBookingStatus()
        End If
    End Sub
    
    Protected Sub btnCancelRequestForSupplies_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        chk_LargeFlyerBags.Checked = False
        chk_SmallFlyerBags.Checked = False
        chk_WindowEnvelopes.Checked = False
        Call ShowCourierBookingStatus()
    End Sub
    
    Protected Sub btnSaveConsignmentChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("vgEditConsignment")
        If Page.IsValid Then
            If SaveConsignmentChanges() Then
                Call BindConsAwaitingCollGrid()
                WebMsgBox.Show("The changes to your consignment have been saved.")
                Call ShowCourierBookingStatus()
            End If
        End If
    End Sub
    
    Protected Sub grd_TodaysCollections_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "cancel" Then
            Dim cell_Collection As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Collection.Text) Then
                plCourierBookingKey = CLng(cell_Collection.Text)
                lblCollectionNumber.Text = plCourierBookingKey
                Call ShowConfirmDeleteCollection()
            End If
        ElseIf e.CommandSource.CommandName = "PrintManifest" Then
        End If
    End Sub
    
    Protected Sub grd_ConsAwaitingCollection_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "edit" Then
            Dim cell_Consignment As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Consignment.Text) Then
                plConsignmentKey = CLng(cell_Consignment.Text)
            End If
            If pbMakeRef1Mandatory = True Then
                lblEditRef1.ForeColor = Drawing.Color.Red
                validator_EditRef1.Enabled = True
            Else
                lblEditRef1.ForeColor = Drawing.Color.Gray
                validator_EditRef1.Enabled = False
            End If
            If pbMakeRef2Mandatory = True Then
                lblEditRef2.ForeColor = Drawing.Color.Red
                validator_EditRef2.Enabled = True
            Else
                lblEditRef2.ForeColor = Drawing.Color.Gray
                validator_EditRef2.Enabled = False
            End If
            If pbMakeRef3Mandatory = True Then
                lblEditRef3.ForeColor = Drawing.Color.Red
                validator_EditRef3.Enabled = True
            Else
                lblEditRef3.ForeColor = Drawing.Color.Gray
                validator_EditRef3.Enabled = False
            End If
            If pbMakeRef4Mandatory = True Then
                lblEditRef4.ForeColor = Drawing.Color.Red
                validator_EditRef4.Enabled = True
            Else
                lblEditRef4.ForeColor = Drawing.Color.Gray
                validator_EditRef4.Enabled = False
            End If
            ResetConsignmentForm()
            GetConsignmentFromKey()
            ShowEditConsignment()
        ElseIf e.CommandSource.CommandName = "collection" Then
            Dim cell_Consignment As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Consignment.Text) Then
                plConsignmentKey = CLng(cell_Consignment.Text)
            End If
            Dim cell_Collection As TableCell = e.Item.Cells (2)
            If IsNumeric(cell_Collection.Text) Then
                plCourierBookingKey = CLng(cell_Collection.Text)
            End If
            If psBookingState = "ZERO" Then
                pbIsCreatingBookingForConsignment = True
                Call ShowCreateAdHocCollection()
            Else
                lblAmendedConsignment.Text = plConsignmentKey
                Call BindAmendedCollectionsGrid()
                Call ShowAmendCollection()
            End If
        ElseIf e.CommandSource.CommandName = "delete" Then
            Dim cell_Consignment As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Consignment.Text) Then
                lblConsignmentNumberToDelete.Text = cell_Consignment.Text
                Call ShowConfirmDeleteConsignment()
            End If
        End If
    End Sub
    
    Protected Sub dgAddressBook_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "select" Then
            Dim cell_Address As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Address.Text) Then
                plCneeAddressKey = CLng(cell_Address.Text)
                Call ResetConsignee()
                Call GetConsigneeAddress()
                Call ShowCreateConsignment()
            End If
        End If
    End Sub
    
    Protected Sub grd_AvailableCollections_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
    End Sub
    
    Protected Sub BindTodaysCollectionsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim sProc As String = "spASPNET_Customer_GetTodaysCourierCollections"
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        If IsNumeric(Session("CustomerKey")) Then
            If plThirdPartyCollectionKey > 0 Then
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = plThirdPartyCollectionKey
            Else
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            End If
        Else
            Server.Transfer("error.aspx ")
        End If
        Try
            oAdapter.Fill(oDataSet, "Collections")
            Dim Source As DataView = oDataSet.Tables("Collections").DefaultView
            If Source.Count > 0 Then
                dgTodaysCollections.DataSource = Source
                dgTodaysCollections.DataBind()
                dgTodaysCollections.Visible = True
                If Source.Count > 5 Then
                    dgTodaysCollections.PagerStyle.Visible = True
                Else
                    dgTodaysCollections.PagerStyle.Visible = False
                End If
                Dim nCollectionCount As Integer = Source.Count
                lblCollectionCount.Text = nCollectionCount.ToString
                If nCollectionCount > 1 Then
                    lblLegendCollection.Text = "collections"
                Else
                    lblLegendCollection.Text = "collection"
                End If
                lblCollectionMessage.Text = ""
                If Source.Count = 1 Then
                    psBookingState = "ONE"
                Else
                    psBookingState = "MANY"
                End If
            Else
                psBookingState = "ZERO"
                lblCollectionCount.Text = "0"
                lblLegendCollection.Text = "collections"
                lblCollectionMessage.Text = "Please book a courier collection at least 1 hour before consignments are ready."
                dgTodaysCollections.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub BindAvailableCollectionsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_GetTodaysCourierCollections", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        If IsNumeric(Session("CustomerKey")) Then
            If plThirdPartyCollectionKey > 0 Then
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = plThirdPartyCollectionKey
            Else
                oAdapter.SelectCommand.Parameters ("@CustomerKey").Value = Session("CustomerKey")
            End If
        Else
            Server.Transfer("error.aspx")
        End If
        Try
            oAdapter.Fill (oDataSet, "AvailableCollections")
            Dim Source As DataView = oDataSet.Tables("AvailableCollections").DefaultView
            If Source.Count > 0 Then
                If Source.Count = 1 Then
                    psBookingState = "ONE"
                    grid_AvailableCollections.Visible = False
                    grid_1AvailableCollection.DataSource = Source
                    grid_1AvailableCollection.DataBind()
                    grid_1AvailableCollection.Visible = True
                    lblAssociatedWithCollection.Text = "Would you like to add this consignment to the collection below?"
                    lblAssocBookingInstructions.Text = "Click 'Yes' to add or 'No' to return to the main status page."
                ElseIf Source.Count > 1 Then
                    psBookingState = "MANY"
                    grid_1AvailableCollection.Visible = False
                    grid_AvailableCollections.DataSource = Source
                    grid_AvailableCollections.DataBind()
                    grid_AvailableCollections.Visible = True
                    lblAssociatedWithCollection.Text = "Would you like to add this consignment to one of the collections below?"
                    lblAssocBookingInstructions.Text = "Click 'Yes' to add your selection or 'No' to return to the main status page."
                End If
            Else
                psBookingState = "ZERO"
                grid_1AvailableCollection.Visible = False
                grid_AvailableCollections.Visible = False
                lblAssociatedWithCollection.Text = "Would you like to book a collection for this consignment now?"
                lblAssocBookingInstructions.Text = "Click 'Yes' to book a collection or 'No' to return to the main status page."
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub BindAmendedCollectionsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_GetTodaysCourierCollections", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        If IsNumeric(Session("CustomerKey")) Then
            If plThirdPartyCollectionKey > 0 Then
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = plThirdPartyCollectionKey
            Else
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            End If
        Else
            Server.Transfer("error.aspx")
        End If
        Try
            oAdapter.Fill(oDataSet, "AvailableCollections")
            Dim Source As DataView = oDataSet.Tables("AvailableCollections").DefaultView
            If Source.Count > 0 Then
                If Source.Count = 1 Then
                    psBookingState = "ONE"
                    grid_ChooseFromCollections.Visible = False
                    grid_OneCollection.DataSource = Source
                    grid_OneCollection.DataBind()
                    grid_OneCollection.Visible = True
                    lblAmendedCollection.Text = "Would you like your consignment to be amended to the collection below?"
                    lblAmendBookingInstructions.Text = "Click 'submit' to add your consignment to the collection above."
                ElseIf Source.Count > 1 Then
                    psBookingState = "MANY"
                    grid_OneCollection.Visible = False
                    grid_ChooseFromCollections.DataSource = Source
                    grid_ChooseFromCollections.DataBind()
                    grid_ChooseFromCollections.Visible = True
                    lblAssociatedWithCollection.Text = "Please select which collection you would like your consignment to be added to"
                    lblAmendBookingInstructions.Text = "Click 'submit' to amend or 'cancel' to return to the main status page."
                End If
            Else
                psBookingState = "ZERO"
                grid_OneCollection.Visible = False
                grid_ChooseFromCollections.Visible = False
                lblAmendedCollection.Text = ""
                lblAmendBookingInstructions.Text = ""
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub BindConsAwaitingCollGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_GetConsAwaitingColl", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        If IsNumeric(Session("CustomerKey")) Then
            oAdapter.SelectCommand.Parameters ("@CustomerKey").Value = Session("CustomerKey")
        Else
            Server.Transfer("error.aspx")
        End If
        Try
            oAdapter.Fill(oDataSet, "Consignments")
            Dim Source As DataView = oDataSet.Tables("Consignments").DefaultView
            If Source.Count > 0 Then
                dgConsignmentsAwaitingCollection.DataSource = Source
                dgConsignmentsAwaitingCollection.DataBind()
                dgConsignmentsAwaitingCollection.Visible = True
            Else
                dgConsignmentsAwaitingCollection.Visible = False
            End If
            Dim nConsignmentCount As Integer = Source.Count
            lblConsignmentCount.Text = nConsignmentCount.ToString
            If nConsignmentCount <> 1 Then
                lblLegendConsignment.Text = "consignments"
            Else
                lblLegendConsignment.Text = "consignment"
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        If pbUseLabelPrinter Then
            dgConsignmentsAwaitingCollection.Columns(8).Visible = False ' Print AWB button
        Else
            dgConsignmentsAwaitingCollection.Columns(9).Visible = False ' print label button
        End If
    End Sub
    
    Protected Sub BindAddressBook()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim sProc As String
        If pbSearchCompanyNameOnly = True Then
            sProc = "spASPNET_UserProfile_GetAddressBookByName"
        Else
            sProc = "spASPNET_UserProfile_GetAddressBook"
        End If
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        Dim sSearchCriteria As String = txtSearchCriteriaAddress.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int ))
            If IsNumeric(Session("UserKey")) Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            Else
                Server.Transfer ("error.aspx")
            End If
            oAdapter.Fill(oDataSet, "Addresses")
            Dim Source As DataView = oDataSet.Tables("Addresses").DefaultView
            If Source.Count > 0 Then
                dgAddressBook.DataSource = Source
                dgAddressBook.DataBind()
                dgAddressBook.Visible = True
                If Source.Count > 12 Then
                    dgAddressBook.PagerStyle.Visible = True
                Else
                    dgAddressBook.PagerStyle.Visible = False
                End If
            Else
                dgAddressBook.Visible = False
                lblAddressMessage.Text = "Nothing found. Please refine your search and try again."
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub dgTodaysCollections_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dgTodaysCollections.CurrentPageIndex = e.NewPageIndex
        Call BindTodaysCollectionsGrid()
    End Sub
    
    Protected Sub dgAddressBook_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dgAddressBook.CurrentPageIndex = e.NewPageIndex
        Call BindAddressBook()
    End Sub
    
    Protected Sub grid_1AvailableCollection_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        grid_1AvailableCollection.CurrentPageIndex = e.NewPageIndex
        Call BindAvailableCollectionsGrid()
    End Sub
    
    Protected Sub grid_AvailableCollections_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        grid_AvailableCollections.CurrentPageIndex = e.NewPageIndex
        Call BindAvailableCollectionsGrid()
    End Sub
    
    Protected Sub CancelCollection(ByVal lCollectionKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As New SqlCommand("spASPNET_CourierBooking_Cancel", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCollectionKey As SqlParameter = New SqlParameter("@CourierBookingKey", SqlDbType.Int, 4)
        oParamCollectionKey.Value = lCollectionKey
        oCmd.Parameters.Add(oParamCollectionKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted , "CancelCollection")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
        Catch ex As SqlException
            oTrans.Rollback("CancelCollection")
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetUsersCompanyDetails()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetAddressDetails", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            oParam.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("CustomerName")) Then
                txtCollectionName.Text = oDataReader("CustomerName")
                Session("CnorName") = oDataReader("CustomerName")
            End If
            If Not IsDBNull(oDataReader("CustomerAddr1")) Then
                txtCollectionAddr1.Text = oDataReader("CustomerAddr1")
                Session("CnorAddr1") = oDataReader("CustomerAddr1")
            End If
            If Not IsDBNull(oDataReader("CustomerAddr2")) Then
                txtCollectionAddr2.Text = oDataReader("CustomerAddr2")
                Session("CnorAddr2") = oDataReader("CustomerAddr2")
            End If
            If Not IsDBNull(oDataReader("CustomerTown")) Then
                txtCollectionCity.Text = oDataReader("CustomerTown")
                Session("CnorCity") = oDataReader("CustomerTown")
            End If
            If Not IsDBNull(oDataReader("CustomerCounty")) Then
                txtCollectionState.Text = oDataReader("CustomerCounty")
                Session("CnorState") = oDataReader("CustomerCounty")
            End If
            If Not IsDBNull(oDataReader("CustomerPostCode")) Then
                txtCollectionPostCode.Text = oDataReader("CustomerPostCode")
                Session("CnorPostCode") = oDataReader("CustomerPostCode")
            End If
            If Not IsDBNull(oDataReader("CountryKey")) Then
                drop_CollectionCountry.SelectedValue = oDataReader("CountryKey")
                plCnorCountryKey = oDataReader("CountryKey")
            End If
            If Not IsDBNull(oDataReader("CountryName")) Then
                Session("CnorCountryName") = oDataReader("CountryName")
            End If
            If Not IsDBNull(oDataReader("Telephone")) Then
                If oDataReader("Telephone") <> "" Then
                    txtCollectionTel.Text = oDataReader("Telephone")
                    txtContactTelephone.Text = oDataReader("Telephone")
                    Session("CnorCtcTel") = oDataReader("Telephone")
                    pbUserHasNoTelOnFile = False
                Else
                    pbUserHasNoTelOnFile = True
                End If
            Else
                pbUserHasNoTelOnFile = True
            End If
            If Not IsDBNull(oDataReader("CollectionPoint")) Then
                If oDataReader("CollectionPoint") <> "" Then
                    txtCollectionPoint.Text = oDataReader("CollectionPoint")
                    pbUserHasNoCollPointOnFile = False
                Else
                    pbUserHasNoCollPointOnFile = True
                End If
            Else
                pbUserHasNoCollPointOnFile = True
            End If
            If Not IsDBNull(oDataReader("EmailAddr")) Then
                psUsersEmailAddr = oDataReader("EmailAddr")
            End If
            oDataReader.Close()
            txtCollectionCTCName.Text = Session("UserName")
            txtContactName.Text = Session("UserName")
            Session("CnorCtcName") = Session("UserName")
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub PopulateCountryDropDowns()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Country_GetCountries", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            Dim oReader2 As SqlDataReader = oCmd.ExecuteReader()
            drop_CneeCountry.DataSource = oReader2
            drop_CneeCountry.DataTextField = "CountryName"
            drop_CneeCountry.DataValueField = "CountryKey"
            drop_CneeCountry.DataBind()
            oReader2.Close()
            Dim oReader3 As SqlDataReader = oCmd.ExecuteReader()
            drop_EditCnorCountry.DataSource = oReader3
            drop_EditCnorCountry.DataTextField = "CountryName"
            drop_EditCnorCountry.DataValueField = "CountryKey"
            drop_EditCnorCountry.DataBind()
            oReader3.Close()
            Dim oReader4 As SqlDataReader = oCmd.ExecuteReader ()
            drop_EditCneeCountry.DataSource = oReader4
            drop_EditCneeCountry.DataTextField = "CountryName"
            drop_EditCneeCountry.DataValueField = "CountryKey"
            drop_EditCneeCountry.DataBind()
            oReader4.Close()
            Dim oReader5 As SqlDataReader = oCmd.ExecuteReader()
            drop_CollectionCountry.DataSource = oReader5
            drop_CollectionCountry.DataTextField = "CountryName"
            drop_CollectionCountry.DataValueField = "CountryKey"
            drop_CollectionCountry.DataBind()
            oReader5.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Function GetCountryIndex(ByVal lCountryKey As Long) As Integer
        Dim x As Integer = 0
        Dim oConn As New SqlConnection(gsConn)
        Dim ds As New DataSet()
        Dim sProc As String = "spASPNET_Country_GetCountries"
        Dim da As New SqlDataAdapter(sProc, oConn)
        da.SelectCommand.CommandType = CommandType.StoredProcedure
        da.Fill(ds)
        Dim dr As DataRow
        For Each dr In ds.Tables(0).Rows
            If lCountryKey = dr(0) Then
                Return x
                Exit For
            End If
            x += 1
        Next
    End Function
    
    Protected Sub GetConsignmentFromKey()
        If plConsignmentKey > 0 Then
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Consignment_GetAWBDetailsFromKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
            oCmd.Parameters.Add(oParam)
            oParam.Value = plConsignmentKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                lblConNoteNumber.Text = plConsignmentKey
                If IsDBNull(oDataReader("CnorName")) Then
                    txtEditCnorName.Text = ""
                Else
                    txtEditCnorName.Text = oDataReader("CnorName")
                End If
                If IsDBNull(oDataReader("CnorAddr1")) Then
                    txtEditCnorAddr1.Text = ""
                Else
                    txtEditCnorAddr1.Text = oDataReader("CnorAddr1")
                End If
                If IsDBNull(oDataReader("CnorAddr2")) Then
                    txtEditCnorAddr2.Text = ""
                Else
                    txtEditCnorAddr2.Text = oDataReader("CnorAddr2")
                End If
                If IsDBNull(oDataReader("CnorTown")) Then
                    txtEditCnorCity.Text = ""
                Else
                    txtEditCnorCity.Text = oDataReader("CnorTown")
                End If
                If IsDBNull(oDataReader("CnorState")) Then
                    txtEditCnorState.Text = ""
                Else
                    txtEditCnorState.Text = oDataReader("CnorState")
                End If
                If IsDBNull(oDataReader("CnorPostCode")) Then
                    txtEditCnorPostCode.Text = ""
                Else
                    txtEditCnorPostCode.Text = oDataReader("CnorPostCode")
                End If
                If Not IsDBNull(oDataReader("CnorCountryKey")) Then
                    drop_EditCnorCountry.SelectedIndex = GetCountryIndex(oDataReader("CnorCountryKey"))
                End If
                If IsDBNull(oDataReader("CnorCtcName")) Then
                    txtEditCnorCtCName.Text = ""
                Else
                    txtEditCnorCtCName.Text = oDataReader("CnorCtcName")
                End If
                If IsDBNull(oDataReader("CnorTel")) Then
                    txtEditCnorTel.Text = ""
                Else
                    txtEditCnorTel.Text = oDataReader("CnorTel")
                End If
    
                If IsDBNull(oDataReader("CneeName")) Then
                    txtEditCneeName.Text = ""
                Else
                    txtEditCneeName.Text = oDataReader("CneeName")
                End If
                If IsDBNull(oDataReader("CneeAddr1")) Then
                    txtEditCneeAddr1.Text = ""
                Else
                    txtEditCneeAddr1.Text = oDataReader("CneeAddr1")
                End If
                If IsDBNull(oDataReader("CneeAddr2")) Then
                    txtEditCneeAddr2.Text = ""
                Else
                    txtEditCneeAddr2.Text = oDataReader("CneeAddr2")
                End If
                If IsDBNull(oDataReader("CneeTown")) Then
                    txtEditCneeCity.Text = ""
                Else
                    txtEditCneeCity.Text = oDataReader("CneeTown")
                End If
                If IsDBNull(oDataReader("CneeState")) Then
                    txtEditCneeState.Text = ""
                Else
                    txtEditCneeState.Text = oDataReader("CneeState")
                End If
                If IsDBNull(oDataReader("CneePostCode")) Then
                    txtEditCneePostCode.Text = ""
                Else
                    txtEditCneePostCode.Text = oDataReader("CneePostCode")
                End If
                If Not IsDBNull(oDataReader("CneeCountryKey")) Then
                    drop_EditCneeCountry.SelectedIndex = GetCountryIndex(oDataReader("CneeCountryKey"))
                End If
                If IsDBNull(oDataReader("CneeCtcName")) Then
                    txtEditCneeCtCName.Text = ""
                Else
                    txtEditCneeCtCName.Text = oDataReader("CneeCtcName")
                End If
                If IsDBNull(oDataReader("CneeTel")) Then
                    txtEditCneeTel.Text = ""
                Else
                    txtEditCneeTel.Text = oDataReader("CneeTel")
                End If
                If IsDBNull(oDataReader("NOP")) Then
                    txtEditNumPieces.Text = ""
                Else
                    txtEditNumPieces.Text = oDataReader("NOP")
                End If
                If IsDBNull(oDataReader("Weight")) Then
                    txtEditWeight.Text = ""
                Else
                    txtEditWeight.Text = oDataReader("Weight")
                End If
                If IsDBNull(oDataReader("ValForCustoms")) Then
                    txtEditValueForCustoms.Text = ""
                Else
                    txtEditValueForCustoms.Text = Format(oDataReader("ValForCustoms"), "#,##0.#0")
                End If
                If Not IsDBNull(oDataReader("ValForCustomsCurKey")) Then
                    Dim nCurrencyCode As Integer = oDataReader("ValForCustomsCurKey")
                    Select Case nCurrencyCode
                        Case 123
                            drop_EditValForCustoms.SelectedIndex = 0
                        Case 52
                            drop_EditValForCustoms.SelectedIndex = 1
                        Case 168
                            drop_EditValForCustoms.SelectedIndex = 2
                        Case Else
                            drop_EditValForCustoms.SelectedIndex = 0
                    End Select
                Else
                    drop_EditValForCustoms.SelectedIndex = 0
                End If
                If IsDBNull(oDataReader("ValForIns")) Then
                    txtEditValueForInsurance.Text = ""
                Else
                    txtEditValueForInsurance.Text = Format(oDataReader("ValForIns"), "#,##0.#0")
                End If
                If Not IsDBNull(oDataReader("ValForInsCurKey")) Then
                    Dim nCurrencyCode As Integer = oDataReader("ValForInsCurKey")
                    Select Case nCurrencyCode
                        Case 123
                            drop_EditValForInsurance.SelectedIndex = 0
                        Case 52
                            drop_EditValForInsurance.SelectedIndex = 1
                        Case 168
                            drop_EditValForInsurance.SelectedIndex = 2
                        Case Else
                            drop_EditValForInsurance.SelectedIndex = 0
                    End Select
                Else
                    drop_EditValForInsurance.SelectedIndex = 0
                End If
                If IsDBNull(oDataReader("Description")) Then
                    txtEditDescription.Text = ""
                Else
                    txtEditDescription.Text = oDataReader("Description")
                End If
                If IsDBNull(oDataReader("SpecialInstructions")) Then
                    txtEditSpclInstructions.Text = ""
                Else
                    txtEditSpclInstructions.Text = oDataReader("SpecialInstructions")
                End If
                If IsDBNull(oDataReader("CustomerRef1")) Then
                    txtEditCustRef1.Text = ""
                Else
                    txtEditCustRef1.Text = oDataReader("CustomerRef1")
                End If
                If IsDBNull(oDataReader("CustomerRef2")) Then
                    txtEditCustRef2.Text = ""
                Else
                    txtEditCustRef2.Text = oDataReader("CustomerRef2")
                End If
                If IsDBNull(oDataReader("Misc1")) Then
                    txtEditCustRef3.Text = ""
                Else
                    txtEditCustRef3.Text = oDataReader("Misc1")
                End If
                If IsDBNull(oDataReader("Misc2")) Then
                    txtEditCustRef4.Text = ""
                Else
                    txtEditCustRef4.Text = oDataReader("Misc2")
                End If
                oDataReader.Close ()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Function SaveConsignmentChanges() As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_Consignment_Update", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
    
        lblEditDateError.Text = ""
        SaveConsignmentChanges = True
    
        Dim paramConsignmentKey As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        paramConsignmentKey.Value = plConsignmentKey
        oCmdAddBooking.Parameters.Add(paramConsignmentKey)
    
        Dim paramCnorName As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        paramCnorName.Value = txtEditCnorName.Text
        oCmdAddBooking.Parameters.Add(paramCnorName)
    
        Dim paramCnorAddr1 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        paramCnorAddr1.Value = txtEditCnorAddr1.Text
        oCmdAddBooking.Parameters.Add(paramCnorAddr1)
    
        Dim paramCnorAddr2 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        paramCnorAddr2.Value = txtEditCnorAddr2.Text
        oCmdAddBooking.Parameters.Add(paramCnorAddr2)
    
        Dim paramCnorTown As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        paramCnorTown.Value = txtEditCnorCity.Text
        oCmdAddBooking.Parameters.Add(paramCnorTown)
    
        Dim paramCnorState As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        paramCnorState.Value = txtEditCnorState.Text
        oCmdAddBooking.Parameters.Add(paramCnorState)
    
        Dim paramCnorPostCode As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        paramCnorPostCode.Value = txtEditCnorPostCode.Text
        oCmdAddBooking.Parameters.Add(paramCnorPostCode)
    
        Dim paramCnorCountryKey As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        paramCnorCountryKey.Value = CLng(drop_EditCnorCountry.SelectedItem.Value)
        oCmdAddBooking.Parameters.Add(paramCnorCountryKey)
    
        Dim paramCnorCtcName As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        paramCnorCtcName.Value = txtEditCnorCtCName.Text
        oCmdAddBooking.Parameters.Add(paramCnorCtcName)
    
        Dim paramCnorTel As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        paramCnorTel.Value = txtEditCnorTel.Text
        oCmdAddBooking.Parameters.Add(paramCnorTel)
    
        Dim paramCneeName As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        paramCneeName.Value = txtEditCneeName.Text
        oCmdAddBooking.Parameters.Add(paramCneeName)
    
        Dim paramCneeAddr1 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        paramCneeAddr1.Value = txtEditCneeAddr1.Text
        oCmdAddBooking.Parameters.Add(paramCneeAddr1)
    
        Dim paramCneeAddr2 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        paramCneeAddr2.Value = txtEditCneeAddr2.Text
        oCmdAddBooking.Parameters.Add(paramCneeAddr2)
    
        Dim paramCneeTown As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        paramCneeTown.Value = txtEditCneeCity.Text
        oCmdAddBooking.Parameters.Add(paramCneeTown)
    
        Dim paramCneeState As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        paramCneeState.Value = txtEditCneeState.Text
        oCmdAddBooking.Parameters.Add(paramCneeState)
    
        Dim paramCneePostCode As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        paramCneePostCode.Value = txtEditCneePostCode.Text
        oCmdAddBooking.Parameters.Add(paramCneePostCode)
    
        Dim paramCneeCountryKey As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        paramCneeCountryKey.Value = CLng(drop_EditCneeCountry.SelectedItem.Value)
        oCmdAddBooking.Parameters.Add(paramCneeCountryKey)
    
        Dim paramCneeCtcName As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        paramCneeCtcName.Value = txtEditCneeCtCName.Text
        oCmdAddBooking.Parameters.Add(paramCneeCtcName)
    
        Dim paramCneeTel As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        paramCneeTel.Value = txtEditCneeTel.Text
        oCmdAddBooking.Parameters.Add(paramCneeTel)
    
        Dim paramNOP As SqlParameter = New SqlParameter("@NOP", SqlDbType.Int , 4)
        paramNOP.Value = CLng(txtEditNumPieces.Text)
        oCmdAddBooking.Parameters.Add(paramNOP)
    
        Dim paramWeight As SqlParameter = New SqlParameter("@Weight", SqlDbType.Money)
        If IsNumeric(txtEditWeight.Text) Then
            paramWeight.Value = CDec(txtEditWeight.Text)
        Else
            paramWeight.Value = 0
        End If
        oCmdAddBooking.Parameters.Add(paramWeight)
    

        Dim paramValForCustoms As SqlParameter = New SqlParameter("@ValForCustoms", SqlDbType.Money)
        If IsNumeric(txtEditValueForCustoms.Text) Then
            paramValForCustoms.Value = CDec(txtEditValueForCustoms.Text)
        Else
            paramValForCustoms.Value = 0
        End If
        oCmdAddBooking.Parameters.Add(paramValForCustoms)
    
        Dim paramValForCustomsCurKey As SqlParameter = New SqlParameter("@ValForCustomsCurrency", SqlDbType.Int, 4)
        If IsNumeric(txtEditValueForCustoms.Text) Then
            paramValForCustomsCurKey.Value = CLng(drop_EditValForCustoms.SelectedItem.Value)
        Else
            paramValForCustomsCurKey.Value = CURRENCY_STERLING
        End If
        oCmdAddBooking.Parameters.Add(paramValForCustomsCurKey)
    
        Dim paramValueForInsurance As SqlParameter = New SqlParameter("@ValForIns", SqlDbType.Money )
        If IsNumeric(txtEditValueForInsurance.Text) Then
            paramValueForInsurance.Value = CDec(txtEditValueForInsurance.Text)
        Else
            paramValueForInsurance.Value = 0
        End If
        oCmdAddBooking.Parameters.Add(paramValueForInsurance)
    
        Dim paramValForInsCur As SqlParameter = New SqlParameter("@ValForInsCurrency", SqlDbType.Int, 4)
        If IsNumeric(txtEditValueForInsurance.Text ) Then
            paramValForInsCur.Value = CLng(drop_EditValForInsurance.SelectedItem.Value)
        Else
            paramValForInsCur.Value = CURRENCY_STERLING
        End If
        oCmdAddBooking.Parameters.Add (paramValForInsCur)

        Dim paramDescription As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        paramDescription.Value = txtEditDescription.Text
        oCmdAddBooking.Parameters.Add (paramDescription)
    
        Dim paramNonDocsFlag As SqlParameter = New SqlParameter("@NonDocsFlag", SqlDbType.Bit, 1)
        If check_EditDocumentsOnly.Checked Then
            paramNonDocsFlag.Value = 0
        Else
            paramNonDocsFlag.Value = 1
        End If
        oCmdAddBooking.Parameters.Add(paramNonDocsFlag)
    
        Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        paramSpecialInstructions.Value = txtEditSpclInstructions.Text
        oCmdAddBooking.Parameters.Add(paramSpecialInstructions)
    
        Dim paramCustomerRef1 As SqlParameter = New SqlParameter("@CustomerRef1", SqlDbType.NVarChar, 30)
        paramCustomerRef1.Value = txtEditCustRef1.Text
        oCmdAddBooking.Parameters.Add(paramCustomerRef1)
    
        Dim paramCustomerRef2 As SqlParameter = New SqlParameter("@CustomerRef2", SqlDbType.NVarChar, 30)
        paramCustomerRef2.Value = txtEditCustRef2.Text
        oCmdAddBooking.Parameters.Add(paramCustomerRef2)
    
        Dim paramCustomerRef3 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramCustomerRef3.Value = txtEditCustRef3.Text
        oCmdAddBooking.Parameters.Add(paramCustomerRef3)
    
        Dim paramCustomerRef4 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        paramCustomerRef4.Value = txtEditCustRef4.Text
        oCmdAddBooking.Parameters.Add(paramCustomerRef4)
    
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction (IsolationLevel.ReadCommitted, "SaveConsignment")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            oTrans.Commit()
            SaveConsignmentChanges = True
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
            oTrans.Rollback("SaveConsignment")
            SaveConsignmentChanges = False
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub GetConsigneeAddress()
        If plCneeAddressKey > 0 Then
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Address_GetAddressFromKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@DestKey", SqlDbType.Int, 4)
            oCmd.Parameters.Add(oParam)
            oParam.Value = plCneeAddressKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                If IsDBNull(oDataReader("Company")) Then
                    txtCneeName.Text = ""
                Else
                    txtCneeName.Text = oDataReader("Company")
                End If
                If IsDBNull(oDataReader("Addr1")) Then
                    txtCneeAddr1.Text = ""
                Else
                    txtCneeAddr1.Text = oDataReader("Addr1")
                End If
                If IsDBNull(oDataReader("Addr2")) Then
                    txtCneeAddr2.Text = ""
                Else
                    txtCneeAddr2.Text = oDataReader("Addr2")
                End If
                If Not IsDBNull(oDataReader("Addr3")) Then
                    txtCneeAddr2.Text &= ", " & oDataReader("Addr3")
                End If
                If IsDBNull(oDataReader("Town")) Then
                    txtCneeCity.Text = ""
                Else
                    txtCneeCity.Text = oDataReader("Town")
                End If
                If IsDBNull(oDataReader("State")) Then
                    txtCneeState.Text = ""
                Else
                    txtCneeState.Text = oDataReader("State")
                End If
                If IsDBNull(oDataReader("PostCode")) Then
                    txtCneePostCode.Text = ""
                Else
                    txtCneePostCode.Text = oDataReader("PostCode")
                End If
                If Not IsDBNull(oDataReader("CountryKey")) Then
                    drop_CneeCountry.SelectedItem.Value = CLng(oDataReader("CountryKey"))
                End If
                If IsDBNull(oDataReader("CountryName")) Then
                    drop_CneeCountry.SelectedItem.Text = ""
                Else
                    drop_CneeCountry.SelectedItem.Text = oDataReader("CountryName")
                End If
                If IsDBNull(oDataReader("AttnOf")) Then
                    txtCneeCTCName.Text = ""
                Else
                    txtCneeCTCName.Text = oDataReader("AttnOf")
                End If
                If IsDBNull(oDataReader("Telephone")) Then
                    txtCneeTel.Text = ""
                Else
                    txtCneeTel.Text = oDataReader("Telephone")
                End If
                'If IsDBNull(oDataReader ("Email")) Then
                '    tbPreAlertEmailAddr01.Text = ""
                'Else
                '    tbPreAlertEmailAddr01.Text = oDataReader ("Email")
                'End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Sub AddCourierBooking()
        'Dim lCourierBookingKey As Long
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_CourierBooking_Add", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
    
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        If IsNumeric(Session("CustomerKey")) Then
            paramCustomerKey.Value = CLng(Session("CustomerKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmdAddBooking.Parameters.Add(paramCustomerKey)
    
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@BookedByKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = CLng(Session("UserKey"))
        oCmdAddBooking.Parameters.Add (paramUserProfileKey)
    
        Dim sDay As String = drop_DayReady.SelectedItem.Text
        Dim sMonth As String = drop_MonthReady.SelectedItem.Text
        Dim sYear As String = drop_YearReady.SelectedItem.Text
        Dim sHour As String = drop_HourReady.SelectedItem.Text
        Dim sMinute As String = drop_MinuteReady.SelectedItem.Text
        Dim sReadyAt As String = sDay & " " & sMonth & " " & sYear & " " & sHour & ":" & sMinute
    
        Dim paramReadyAt As SqlParameter = New SqlParameter("@ReadyAt", SqlDbType.DateTime)
        paramReadyAt.Value = CDate(sReadyAt)
        oCmdAddBooking.Parameters.Add(paramReadyAt)
    
        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = txtCollectionName.Text
        oCmdAddBooking.Parameters.Add(paramCompany)
    
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = txtCollectionAddr1.Text
        oCmdAddBooking.Parameters.Add(paramAddr1)
    
        Dim paramAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramAddr2.Value = txtCollectionAddr2.Text
        oCmdAddBooking.Parameters.Add(paramAddr2)
    
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = txtCollectionCity.Text
        oCmdAddBooking.Parameters.Add(paramTown)
    
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        paramState.Value = txtCollectionState.Text
        oCmdAddBooking.Parameters.Add(paramState)
    
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = txtCollectionPostCode.Text
        oCmdAddBooking.Parameters.Add(paramPostCode)
    
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = CLng(drop_CollectionCountry.SelectedItem.Value)
        oCmdAddBooking.Parameters.Add(paramCountryKey)
    
        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.NVarChar, 50)
        paramContactName.Value = txtCollectionCTCName.Text
        oCmdAddBooking.Parameters.Add(paramContactName)
    
        Dim paramContactTelNo As SqlParameter = New SqlParameter("@ContactTelNo", SqlDbType.NVarChar, 50)
        paramContactTelNo.Value = txtCollectionTel.Text
        oCmdAddBooking.Parameters.Add(paramContactTelNo)
    
        Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 50)
        paramCollectionPoint.Value = txtCollectionPoint.Text
        oCmdAddBooking.Parameters.Add(paramCollectionPoint)
    
        Dim paramCollectionInfo As SqlParameter = New SqlParameter("@CollectionInfo", SqlDbType.NVarChar, 1000)
        paramCollectionInfo.Value = txtNoteToDriver.Text
        oCmdAddBooking.Parameters.Add(paramCollectionInfo)
    
        Dim paramCourierBookingKey As SqlParameter = New SqlParameter("@CourierBookingKey", SqlDbType.Int, 4)
        paramCourierBookingKey.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(paramCourierBookingKey)
    
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddCourierBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery ()
            plCourierBookingKey = CLng(oCmdAddBooking.Parameters("@CourierBookingKey").Value.ToString)
            lblCourierBookingNumber.Text = plCourierBookingKey
            oTrans.Commit()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
            oTrans.Rollback("AddCourierBooking")
        Finally
            oConn.Close()
        End Try
    
        If pbUserHasNoTelOnFile Then
            SaveUsersTelephoneNumber()
        End If
        If pbUserHasNoCollPointOnFile Then
            SaveUsersCollectionPoint()
        End If
        If pbIsCreatingBookingForConsignment Then
            AssociateConsignmentWithCollection(plConsignmentKey, plCourierBookingKey)
        End If
        'reset flags
        pbIsCreatingBookingForConsignment = False
        BindTodaysCollectionsGrid()
    End Sub
    
    Protected Sub SaveUsersTelephoneNumber()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddTelephone As SqlCommand = New SqlCommand("spASPNET_UserProfile_AddTelephone", oConn)
        oCmdAddTelephone.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            paramUserProfileKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmdAddTelephone.Parameters.Add(paramUserProfileKey)
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
        paramTelephone.Value = txtCollectionTel.Text
        oCmdAddTelephone.Parameters.Add(paramTelephone)
        Try
            oConn.Open()
            oCmdAddTelephone.Connection = oConn
            oCmdAddTelephone.ExecuteNonQuery ()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SaveUsersCollectionPoint()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdCollectionPoint As SqlCommand = New SqlCommand("spASPNET_UserProfile_AddCollectionPoint", oConn)
        oCmdCollectionPoint.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            paramUserProfileKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmdCollectionPoint.Parameters.Add(paramUserProfileKey)
        Dim paramCollectionPoint As SqlParameter = New SqlParameter("@CollectionPoint", SqlDbType.NVarChar, 520)
        paramCollectionPoint.Value = txtCollectionPoint.Text
        oCmdCollectionPoint.Parameters.Add(paramCollectionPoint)
        Try
            oConn.Open()
            oCmdCollectionPoint.Connection = oConn
            oCmdCollectionPoint.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close ()
        End Try
    End Sub
    
    Protected Sub AssociateConsignmentWithCollection(ByVal lConsignmentKey As Long, ByVal lCourierBookingKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddTelephone As SqlCommand = New SqlCommand("spASPNET_Consignment_AssociateWithCollection", oConn)
        oCmdAddTelephone.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            paramUserProfileKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmdAddTelephone.Parameters.Add(paramUserProfileKey)
        Dim paramConsignmentKey As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        paramConsignmentKey.Value = lConsignmentKey
        oCmdAddTelephone.Parameters.Add(paramConsignmentKey)
        Dim paramCourierBookingKey As SqlParameter = New SqlParameter("@CourierBookingKey", SqlDbType.Int, 4)
        paramCourierBookingKey.Value = lCourierBookingKey
        oCmdAddTelephone.Parameters.Add(paramCourierBookingKey)
        Try
            oConn.Open()
            oCmdAddTelephone.Connection = oConn
            oCmdAddTelephone.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close ()
        End Try
    End Sub
    
    Protected Sub UnBookCollection(ByVal lConsignmentKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Consignment_UnAssociateWithCollection", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            paramUserProfileKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramConsignmentKey As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        paramConsignmentKey.Value = lConsignmentKey
        oCmd.Parameters.Add(paramConsignmentKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function AddNewConsignment() As Long
        If IsValid Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oTrans As SqlTransaction
            Dim oCmdAddConsignment As SqlCommand = New SqlCommand("spASPNET_Consignment_AddNew", oConn)
            oCmdAddConsignment.CommandType = CommandType.StoredProcedure
    
            lblError.Text = ""
    
            'Put the ReadyOn datetime together
            Dim sDay As String = drop_DayReady.SelectedItem.Text
            Dim sMonth As String = drop_MonthReady.SelectedItem.Text
            Dim sYear As String = drop_YearReady.SelectedItem.Text
            Dim sHour As String = drop_HourReady.SelectedItem.Text
            Dim sMinute As String = drop_MinuteReady.SelectedItem.Text
            Dim sReadyAt As String
    
            Try
                sReadyAt = DateTime.Parse(sDay & " " & sMonth & " " & sYear & " " & sHour & ":" & sMinute)
            Catch ex As Exception
                lblError.Text = "Invalid collection date/time"
                Exit Function
            End Try
    
            Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
            paramUserProfileKey.Value = CLng(Session("UserKey"))
            oCmdAddConsignment.Parameters.Add(paramUserProfileKey)
    
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            If IsNumeric(Session("CustomerKey")) Then
                paramCustomerKey.Value = CLng(Session("CustomerKey"))
            Else
                Server.Transfer ("error.aspx")
            End If
            oCmdAddConsignment.Parameters.Add(paramCustomerKey)
    
            Dim paramConsignmentType As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
            paramConsignmentType.Value = "EXPRESS"
            oCmdAddConsignment.Parameters.Add(paramConsignmentType)
    
            Dim paramConsignmentService As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 8)
            paramConsignmentService.Value = Nothing 'drop_ServiceLevel.SelectedItem.Text
            oCmdAddConsignment.Parameters.Add(paramConsignmentService)
    
            Dim paramValForCustoms As SqlParameter = New SqlParameter("@ValueForCustoms", SqlDbType.Money)
            If IsNumeric(txtValForCustoms.Text) Then
                paramValForCustoms.Value = CDec(txtValForCustoms.Text)
            Else
                paramValForCustoms.Value = 0
            End If
            oCmdAddConsignment.Parameters.Add(paramValForCustoms)
    
            Dim paramValForCustomsCurKey As SqlParameter = New SqlParameter("@ValForCustomsCurKey", SqlDbType.Int, 4)
            If IsNumeric( txtValForCustoms.Text) Then
                paramValForCustomsCurKey.Value = CLng(drop_ValForInsurance.SelectedItem.Value)
            Else
                paramValForCustomsCurKey.Value = CURRENCY_STERLING
            End If
            oCmdAddConsignment.Parameters.Add(paramValForCustomsCurKey)
    
            Dim paramValueForInsurance As SqlParameter = New SqlParameter("@ValueForInsurance", SqlDbType.Money)
            If IsNumeric( txtValForInsurance.Text) Then
                paramValueForInsurance.Value = CDec(txtValForInsurance.Text)
            Else
                paramValueForInsurance.Value = 0
            End If
            oCmdAddConsignment.Parameters.Add (paramValueForInsurance)
    
            Dim paramValForInsCurKey As SqlParameter = New SqlParameter("@ValForInsCurKey", SqlDbType.Int, 4)
            If IsNumeric(txtValForInsurance.Text) Then
                paramValForInsCurKey.Value = CLng(drop_ValForInsurance.SelectedItem.Value)
            Else
                paramValForInsCurKey.Value = CURRENCY_STERLING
            End If
            oCmdAddConsignment.Parameters.Add (paramValForInsCurKey)
    
            Dim paramDescription As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
            paramDescription.Value = txtDescription.Text
            oCmdAddConsignment.Parameters.Add(paramDescription)
    
            Dim paramNonDocsFlag As SqlParameter = New SqlParameter("@NonDocsFlag", SqlDbType.Bit, 1)
            If check_DocumentsOnly.Checked Then
                paramNonDocsFlag.Value = 0
            Else
                paramNonDocsFlag.Value = 1
            End If
            oCmdAddConsignment.Parameters.Add(paramNonDocsFlag)
    
            Dim paramNumberOfPieces As SqlParameter = New SqlParameter("@NumberOfPieces", SqlDbType.Int, 4)
            If IsNumeric(txtNoPieces.Text) Then
                paramNumberOfPieces.Value = CInt(txtNoPieces.Text)
            Else
                paramNumberOfPieces.Value = 0
            End If
            oCmdAddConsignment.Parameters.Add(paramNumberOfPieces)
    
            Dim paramWeight As SqlParameter = New SqlParameter("@Weight", SqlDbType.Decimal)
            If IsNumeric(txtWeight.Text ) Then
                paramWeight.Value = CDec(txtWeight.Text)
            Else
                paramWeight.Value = 0
            End If
            oCmdAddConsignment.Parameters.Add(paramWeight)
    
            Dim paramShippersRef1 As SqlParameter = New SqlParameter("@ShippersRef1", SqlDbType.NVarChar, 25)
            paramShippersRef1.Value = txtCustRef1.Text
            oCmdAddConsignment.Parameters.Add (paramShippersRef1)
    
            Dim paramShippersRef2 As SqlParameter = New SqlParameter("@ShippersRef2", SqlDbType.NVarChar, 25)
            paramShippersRef2.Value = txtCustRef2.Text
            oCmdAddConsignment.Parameters.Add(paramShippersRef2)
    
            Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
            paramMisc1.Value = txtCustRef3.Text
            oCmdAddConsignment.Parameters.Add(paramMisc1)
    
            Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
            paramMisc2.Value = txtCustRef4.Text
            oCmdAddConsignment.Parameters.Add(paramMisc2)
    
            Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
            paramSpecialInstructions.Value = txtSpecialInstructions.Text
            oCmdAddConsignment.Parameters.Add(paramSpecialInstructions)
    
            Dim paramTimedDeliveryFlag As SqlParameter = New SqlParameter("@TimedDeliveryFlag", SqlDbType.Bit, 1)
            If txtSpecialInstructions.Text <> "" Then
                paramTimedDeliveryFlag.Value = 1
            Else
                paramTimedDeliveryFlag.Value = 0
            End If
            oCmdAddConsignment.Parameters.Add(paramTimedDeliveryFlag)
    
            Dim paramCnorName As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
            paramCnorName.Value = Session("CnorName")
            oCmdAddConsignment.Parameters.Add(paramCnorName)
    
            Dim paramCnorAddr1 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
            paramCnorAddr1.Value = Session("CnorAddr1")
            oCmdAddConsignment.Parameters.Add(paramCnorAddr1)
    
            Dim paramCnorAddr2 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
            paramCnorAddr2.Value = Session("CnorAddr2")
            oCmdAddConsignment.Parameters.Add(paramCnorAddr2)
    
            Dim paramCnorTown As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
            paramCnorTown.Value = Session("CnorCity")
            oCmdAddConsignment.Parameters.Add(paramCnorTown)
    
            Dim paramCnorState As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
            paramCnorState.Value = Session("CnorState")
            oCmdAddConsignment.Parameters.Add(paramCnorState)
    
            Dim paramCnorPostCode As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
            paramCnorPostCode.Value = Session("CnorPostCode")
            oCmdAddConsignment.Parameters.Add(paramCnorPostCode)
    
            Dim paramCnorCountryKey As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
            paramCnorCountryKey.Value = plCnorCountryKey
            oCmdAddConsignment.Parameters.Add(paramCnorCountryKey)
    
            Dim paramCnorCtcName As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
            paramCnorCtcName.Value = txtCollectionCTCName.Text
            oCmdAddConsignment.Parameters.Add(paramCnorCtcName)
    
            Dim paramCnorTel As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
            paramCnorTel.Value = txtCollectionTel.Text
            oCmdAddConsignment.Parameters.Add(paramCnorTel)
    
            Dim paramCnorEmail As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 100)
            paramCnorEmail.Value = psUsersEmailAddr
            oCmdAddConsignment.Parameters.Add(paramCnorEmail)
    
            Dim paramCneeName As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
            paramCneeName.Value = txtCneeName.Text
            oCmdAddConsignment.Parameters.Add(paramCneeName)
    
            Dim paramCneeAddr1 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
            paramCneeAddr1.Value = txtCneeAddr1.Text
            oCmdAddConsignment.Parameters.Add(paramCneeAddr1)
    
            Dim paramCneeAddr2 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
            paramCneeAddr2.Value = txtCneeAddr2.Text
            oCmdAddConsignment.Parameters.Add(paramCneeAddr2)
    
            Dim paramCneeTown As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
            paramCneeTown.Value = txtCneeCity.Text
            oCmdAddConsignment.Parameters.Add(paramCneeTown)
    
            Dim paramCneeState As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
            paramCneeState.Value = txtCneeState.Text
            oCmdAddConsignment.Parameters.Add(paramCneeState)
    
            Dim paramCneePostCode As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
            paramCneePostCode.Value = txtCneePostCode.Text
            oCmdAddConsignment.Parameters.Add(paramCneePostCode)
    
            Dim paramCneeCountryKey As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
            paramCneeCountryKey.Value = CLng(drop_CneeCountry.SelectedItem.Value)
            oCmdAddConsignment.Parameters.Add(paramCneeCountryKey)
    
            Dim paramCneeCtcName As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
            paramCneeCtcName.Value = txtCneeCTCName.Text
            oCmdAddConsignment.Parameters.Add(paramCneeCtcName)
    
            Dim paramCneeTel As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
            paramCneeTel.Value = txtCneeTel.Text
            oCmdAddConsignment.Parameters.Add(paramCneeTel)
    
            Dim paramCneeEmail As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 100)
            paramCneeEmail.Value = txtCneeEmail.Text
            oCmdAddConsignment.Parameters.Add(paramCneeEmail)
    
            Dim paramConsignmentKey As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
            paramConsignmentKey.Direction = ParameterDirection.Output
            oCmdAddConsignment.Parameters.Add(paramConsignmentKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddConsignment")
                oCmdAddConsignment.Connection = oConn
                oCmdAddConsignment.Transaction = oTrans
                oCmdAddConsignment.ExecuteNonQuery ()
                plConsignmentKey = CLng(oCmdAddConsignment.Parameters("@ConsignmentKey").Value.ToString)
                lblConsignmentNumber.Text = plConsignmentKey
    
                Dim oCmdAddConsignmentTracking As SqlCommand = New SqlCommand("spASPNET_Consignment_AddTrackingRequests3", oConn)
                oCmdAddConsignmentTracking.CommandType = CommandType.StoredProcedure
                Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
                param1.Value = CLng(Session("UserKey"))
                oCmdAddConsignmentTracking.Parameters.Add (param1)
                Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                If IsNumeric(Session("CustomerKey")) Then
                    param2.Value = CLng(Session("CustomerKey"))
                Else
                    oTrans.Rollback("AddConsignment")
                    oConn.Close()
                    Server.Transfer("error.aspx ")
                End If
                oCmdAddConsignmentTracking.Parameters.Add(param2)
                Dim param3 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
                param3.Value = plConsignmentKey
                oCmdAddConsignmentTracking.Parameters.Add(param3)
                Dim param4 As SqlParameter = New SqlParameter("@SMSTracking_MobileNo", SqlDbType.NVarChar, 20)
                If chk_SMSTextTracking.Checked = True Then
                    If IsNumeric(txtTrackingMobileNo.Text) Then
                        param4.Value = txtTrackingMobileNo.Text
                    Else
                        param4.Value = Nothing
                    End If
                Else
                    param4.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add (param4)
                Dim param5 As SqlParameter = New SqlParameter("@EmailTracking_EmailAddr", SqlDbType.NVarChar, 100)
                If chk_EmailTracking.Checked = True Then
                    If txtTrackingEmailAddr.Text <> "" Then
                        param5.Value = txtTrackingEmailAddr.Text
                    Else
                        param5.Value = Nothing
                    End If
                Else
                    param5.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(param5)
                
                Dim paramEmailPreAlert_EmailAddr01 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr01", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert01.Checked = True Then
                    If tbPreAlertEmailAddr01.Text <> "" Then
                        paramEmailPreAlert_EmailAddr01.Value = tbPreAlertEmailAddr01.Text
                    Else
                        paramEmailPreAlert_EmailAddr01.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr01.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr01)

                Dim paramEmailDeliveryConfirmation_EmailAddr01 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr01", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation01.Checked = True Then
                    If tbPreAlertEmailAddr01.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr01.Value = tbPreAlertEmailAddr01.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr01.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr01.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr01)

                Dim paramEmailPreAlert_EmailAddr02 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr02", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert02.Checked = True Then
                    If tbPreAlertEmailAddr02.Text <> "" Then
                        paramEmailPreAlert_EmailAddr02.Value = tbPreAlertEmailAddr01.Text
                    Else
                        paramEmailPreAlert_EmailAddr02.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr02.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr02)

                Dim paramEmailDeliveryConfirmation_EmailAddr02 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr02", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation02.Checked = True Then
                    If tbPreAlertEmailAddr01.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr02.Value = tbPreAlertEmailAddr02.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr02.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr02.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr02)

                Dim paramEmailPreAlert_EmailAddr03 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr03", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert03.Checked = True Then
                    If tbPreAlertEmailAddr03.Text <> "" Then
                        paramEmailPreAlert_EmailAddr03.Value = tbPreAlertEmailAddr03.Text
                    Else
                        paramEmailPreAlert_EmailAddr03.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr03.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr03)

                Dim paramEmailDeliveryConfirmation_EmailAddr03 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr03", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation03.Checked = True Then
                    If tbPreAlertEmailAddr01.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr03.Value = tbPreAlertEmailAddr03.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr03.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr03.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr03)

                Dim paramEmailPreAlert_EmailAddr04 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr04", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert04.Checked = True Then
                    If tbPreAlertEmailAddr04.Text <> "" Then
                        paramEmailPreAlert_EmailAddr04.Value = tbPreAlertEmailAddr04.Text
                    Else
                        paramEmailPreAlert_EmailAddr04.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr04.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr04)

                Dim paramEmailDeliveryConfirmation_EmailAddr04 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr04", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation04.Checked = True Then
                    If tbPreAlertEmailAddr04.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr04.Value = tbPreAlertEmailAddr04.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr04.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr04.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr04)

                Dim paramEmailPreAlert_EmailAddr05 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr05", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert05.Checked = True Then
                    If tbPreAlertEmailAddr05.Text <> "" Then
                        paramEmailPreAlert_EmailAddr05.Value = tbPreAlertEmailAddr05.Text
                    Else
                        paramEmailPreAlert_EmailAddr05.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr05.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr05)

                Dim paramEmailDeliveryConfirmation_EmailAddr05 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr05", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation05.Checked = True Then
                    If tbPreAlertEmailAddr05.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr05.Value = tbPreAlertEmailAddr05.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr05.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr05.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr05)

                Dim paramEmailPreAlert_EmailAddr06 As SqlParameter = New SqlParameter("@EmailPreAlert_EmailAddr06", SqlDbType.NVarChar, 100)
                If cbEmailPreAlert06.Checked = True Then
                    If tbPreAlertEmailAddr06.Text <> "" Then
                        paramEmailPreAlert_EmailAddr06.Value = tbPreAlertEmailAddr06.Text
                    Else
                        paramEmailPreAlert_EmailAddr06.Value = Nothing
                    End If
                Else
                    paramEmailPreAlert_EmailAddr06.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailPreAlert_EmailAddr06)

                Dim paramEmailDeliveryConfirmation_EmailAddr06 As SqlParameter = New SqlParameter("@EmailDeliveryConfirmation_EmailAddr06", SqlDbType.NVarChar, 100)
                If cbEmailDeliveryConfirmation06.Checked = True Then
                    If tbPreAlertEmailAddr06.Text <> "" Then
                        paramEmailDeliveryConfirmation_EmailAddr06.Value = tbPreAlertEmailAddr06.Text
                    Else
                        paramEmailDeliveryConfirmation_EmailAddr06.Value = Nothing
                    End If
                Else
                    paramEmailDeliveryConfirmation_EmailAddr06.Value = Nothing
                End If
                oCmdAddConsignmentTracking.Parameters.Add(paramEmailDeliveryConfirmation_EmailAddr06)

                oCmdAddConsignmentTracking.Connection = oConn
                oCmdAddConsignmentTracking.Transaction = oTrans
                oCmdAddConsignmentTracking.ExecuteNonQuery ()
    
                oTrans.Commit()
    
                chk_SMSTextTracking.Checked = False
                chk_EmailTracking.Checked = False
                cbEmailPreAlert01.Checked = False
    
            Catch ex As SqlException
                lblError.Text = ""
                lblError.Text = ex.ToString
                oTrans.Rollback ("AddConsignment")
            Finally
                pbIsCreatingNewConsignment = False
                oConn.Close()
            End Try
    
            If pbUserHasNoTelOnFile Then
                SaveUsersTelephoneNumber()
            End If
            If pbUserHasNoCollPointOnFile Then
                SaveUsersCollectionPoint()
            End If
        End If
        Return plConsignmentKey
    End Function
    
    Protected Sub DeleteConsignment(ByVal lConsignmentKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Consignment_Delete", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
    
        Dim oParamUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            oParamUserProfileKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmd.Parameters.Add(oParamUserProfileKey)
    
        Dim paramConsignmentKey As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        paramConsignmentKey.Value = lConsignmentKey
        oCmd.Parameters.Add(paramConsignmentKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub AddRequestForSupplies(ByVal sProduct As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddRequestForSupplies As SqlCommand = New SqlCommand("spASPNET_RequestedSupplies_Add", oConn)
        oCmdAddRequestForSupplies.CommandType = CommandType.StoredProcedure
    
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        If IsNumeric(Session("UserKey")) Then
            paramUserKey.Value = CLng(Session("UserKey"))
        Else
            Server.Transfer("error.aspx")
        End If
        oCmdAddRequestForSupplies.Parameters.Add(paramUserKey)
    
        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.NVarChar, 50)
        paramContactName.Value = txtContactName.Text
        oCmdAddRequestForSupplies.Parameters.Add(paramContactName)
    
        Dim paramContactEmail As SqlParameter = New SqlParameter("@ContactEmail", SqlDbType.NVarChar, 100)
        paramContactEmail.Value = txtContactEmail.Text
        oCmdAddRequestForSupplies.Parameters.Add(paramContactEmail)
    
        Dim paramContactTelephone As SqlParameter = New SqlParameter("@ContactTelephone", SqlDbType.NVarChar, 50)
        paramContactTelephone.Value = txtContactTelephone.Text
        oCmdAddRequestForSupplies.Parameters.Add(paramContactTelephone)
    
        Dim paramProduct As SqlParameter = New SqlParameter("@Product", SqlDbType.NVarChar, 50)
        paramProduct.Value = sProduct
        oCmdAddRequestForSupplies.Parameters.Add(paramProduct)
    
        Try
            oConn.Open()
            oCmdAddRequestForSupplies.Connection = oConn
            oCmdAddRequestForSupplies.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub AddAddressToPersonalAddressBook()
        Dim bError As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_Add", oConn)
        Dim oTrans As SqlTransaction
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        If IsNumeric(Session("CustomerKey")) Then
            paramCustomerKey.Value = Session("CustomerKey")
        Else
            Server.Transfer("error.aspx")
        End If
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramCode As SqlParameter = New SqlParameter("@Code", SqlDbType.NVarChar, 20)
        paramCode.Value = Nothing
        oCmd.Parameters.Add(paramCode)
        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = txtCneeName.Text
        oCmd.Parameters.Add(paramCompany)
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = txtCneeAddr1.Text
        oCmd.Parameters.Add(paramAddr1)
        Dim paramparamAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramparamAddr2.Value = txtCneeAddr2.Text
        oCmd.Parameters.Add (paramparamAddr2)
        Dim paramparamAddr3 As SqlParameter = New SqlParameter("@Addr3", SqlDbType.NVarChar, 50)
        paramparamAddr3.Value = Nothing
        oCmd.Parameters.Add(paramparamAddr3)
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = txtCneeCity.Text
        oCmd.Parameters.Add(paramTown)
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        paramState.Value = txtCneeState.Text
        oCmd.Parameters.Add(paramState)
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = txtCneePostCode.Text
        oCmd.Parameters.Add(paramPostCode)
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = CLng(drop_CneeCountry.SelectedItem.Value)
        oCmd.Parameters.Add(paramCountryKey)
        Dim paramDefaultCommodityId As SqlParameter = New SqlParameter("@DefaultCommodityId", SqlDbType.NVarChar, 100)
        paramDefaultCommodityId.Value = Nothing
        oCmd.Parameters.Add(paramDefaultCommodityId)
        Dim paramDefaultSpecialInstructions As SqlParameter = New SqlParameter("@DefaultSpecialInstructions", SqlDbType.NVarChar, 100)
        paramDefaultSpecialInstructions.Value = Nothing
        oCmd.Parameters.Add(paramDefaultSpecialInstructions)
        Dim paramAttnOf As SqlParameter = New SqlParameter("@AttnOf", SqlDbType.NVarChar, 50)
        paramAttnOf.Value = txtCneeCTCName.Text
        oCmd.Parameters.Add(paramAttnOf)
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar , 50)
        paramTelephone.Value = txtCneeTel.Text
        oCmd.Parameters.Add(paramTelephone)
        Dim paramFax As SqlParameter = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        paramFax.Value = Nothing
        oCmd.Parameters.Add(paramFax)
        Dim paramEmail As SqlParameter = New SqlParameter("@Email", SqlDbType.NVarChar, 50)
        paramEmail.Value = Nothing
        oCmd.Parameters.Add (paramEmail)
        Dim paramLastUpdatedByKey As SqlParameter = New SqlParameter("@LastUpdatedByKey", SqlDbType.Int, 4)
        paramLastUpdatedByKey.Value = Session("UserKey")
        oCmd.Parameters.Add (paramLastUpdatedByKey)
        Dim paramAddressKey As SqlParameter = New SqlParameter("@AddressKey", SqlDbType.Int, 4)     'OUTPUT PARAMETER
        paramAddressKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramAddressKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            plCneeAddressKey = paramAddressKey.Value
        Catch ex As SqlException
            oTrans.Rollback("AddRecord")
            lblError.Text = ex.ToString
            oConn.Close()
            bError = True
        End Try
        If plCneeAddressKey > 0 Then
            Dim oCmd2 As SqlCommand = New SqlCommand("spASPNET_Address_AddToPersonal", oConn)
            oCmd2.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = Session("UserKey")
            oCmd2.Parameters.Add(paramUserKey)
            Dim paramGABKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramGABKey.Value = plCneeAddressKey
            oCmd2.Parameters.Add(paramGABKey)
            Try
                oCmd2.Connection = oConn
                oCmd2.Transaction = oTrans
                oCmd2.ExecuteNonQuery()
                oTrans.Commit ()
            Catch ex As SqlException
                oTrans.Rollback("AddRecord")
                lblError.Text = ex.ToString
                bError = True
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Sub ResetConsignee()
        txtCneeName.Text = ""
        txtCneeAddr1.Text = ""
        txtCneeAddr2.Text = ""
        txtCneeCity.Text = ""
        txtCneeState.Text = ""
        txtCneePostCode.Text = ""
        txtCneeCTCName.Text = ""
        txtCneeTel.Text = ""
        drop_CneeCountry.SelectedIndex = -1
        drop_CneeCountry.SelectedItem.Text = "- please select -"
    End Sub
    
    Protected Sub ResetConsignmentForm()
        txtValForInsurance.Text = ""
        txtValForCustoms.Text = ""
        txtWeight.Text = ""
        txtSpecialInstructions.Text = ""
        txtCustRef1.Text = ""
        txtCustRef2.Text = ""
        txtCustRef3.Text = ""
        txtCustRef4.Text = ""
        chk_SaveAddress.Checked = False
        Call ResetConsignee()
    End Sub

    Protected Sub CheckValidNumber_ServerValidate(ByVal source As Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs)
        Dim oCustomValidator As CustomValidator = source
        Dim sName As String = oCustomValidator.ControlToValidate
        If IsNumeric(args.Value) AndAlso CDec( args.Value) >= 0 Then
            args.IsValid = True
        Else
            args.IsValid = False
        End If
    End Sub
    
    Protected Sub btnCreateConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CreateConsignment()
    End Sub
    
    Protected Sub btnBookACollection_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCreateAdHocCollection()
    End Sub

    Protected Sub btnRequestSupplies_Click_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowRequestSupplies()
    End Sub
    
    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SearchAddresses()
    End Sub
    
    Protected Sub btnShowAllAddresses_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllAddresses()
    End Sub
    
    Protected Sub btnGoBackToConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GoBackToConsignment()
    End Sub
    
    Protected Sub btnShowCourierBookingStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisplayCourierBookingStatus()
    End Sub
    
    Protected Sub ClearPreAlertAddresses()
        trPreAlert02.Visible = False
        trPreAlert03.Visible = False
        trPreAlert04.Visible = False
        trPreAlert05.Visible = False
        trPreAlert06.Visible = False
        lnkbtnAddAnotherPreAlert.Visible = True
    End Sub
    
    Protected Sub btnCancelConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs )
        Call CancelConsignment()
    End Sub
    
    Protected Sub btnSubmitNewConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SubmitNewConsignment()
    End Sub
    
    Protected Sub btnSubmitNewCollection_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SubmitCourierCollection()
    End Sub
    
    Protected Sub btnCancelCollection_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisplayCourierBookingStatus()
    End Sub
    
    Protected Sub lnkbtnAddAnotherPreAlert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If trPreAlert02.Visible = False Then
            trPreAlert02.Visible = True
            Exit Sub
        End If
        If trPreAlert03.Visible = False Then
            trPreAlert03.Visible = True
            Exit Sub
        End If
        If trPreAlert04.Visible = False Then
            trPreAlert04.Visible = True
            Exit Sub
        End If
        If trPreAlert05.Visible = False Then
            trPreAlert05.Visible = True
            Exit Sub
        End If
        If trPreAlert06.Visible = False Then
            trPreAlert06.Visible = True
            lnkbtnAddAnotherPreAlert.Visible = False
            Exit Sub
        End If
    End Sub
    
    Function sDriversManifestArgs(ByVal DataItem As Object) As String
        sDriversManifestArgs = DataItem("Key")
    End Function
    
    Property plConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("ConsignmentKey") = Value
        End Set
    End Property
    
    Property plCourierBookingKey() As Long
        Get
            Dim o As Object = ViewState("CourierBookingKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CourierBookingKey") = Value
        End Set
    End Property
    
    Property plCneeAddressKey() As Long
        Get
            Dim o As Object = ViewState("CneeAddressKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CneeAddressKey") = Value
        End Set
    End Property
    
    Property plCnorCountryKey() As Long
        Get
            Dim o As Object = ViewState("CnorCountryKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CnorCountryKey") = Value
        End Set
    End Property
    
    Property pbUserHasNoTelOnFile() As Boolean
        Get
            Dim o As Object = ViewState("UserHasNoTelOnFile")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UserHasNoTelOnFile") = Value
        End Set
    End Property
    
    Property pbUserHasNoCollPointOnFile() As Boolean
        Get
            Dim o As Object = ViewState("UserHasNoCollPointOnFile")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UserHasNoCollPointOnFile") = Value
        End Set
    End Property
    
    Property psUsersEmailAddr() As String
        Get
            Dim o As Object = ViewState("UsersEmailAddr")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("UsersEmailAddr") = Value
            txtTrackingEmailAddr.Text = value
            tbPreAlertEmailAddr01.Text = value
            txtContactEmail.Text = value
        End Set
    End Property
    
    Property pbIsCreatingBookingForConsignment() As Boolean
        Get
            Dim o As Object = ViewState("IsCreatingBookingForConsignment")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IsCreatingBookingForConsignment") = Value
        End Set
    End Property
    
    Property pbIsAssociatingConsignmentWithBooking() As Boolean
        Get
            Dim o As Object = ViewState("IsAssociatingConsignmentWithBooking")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IsAssociatingConsignmentWithBooking") = Value
        End Set
    End Property
    
    Property pbUseLabelPrinter() As Boolean
        Get
            Dim o As Object = ViewState("UseLabelPrinter")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("UseLabelPrinter") = Value
        End Set
    End Property
    
    Property pbIsCreatingNewConsignment() As Boolean
        Get
            Dim o As Object = ViewState("IsCreatingNewConsignment")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IsCreatingNewConsignment") = Value
        End Set
    End Property
    
    Property psBookingState() As String
        Get
            Dim o As Object = ViewState("BookingState")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("BookingState") = Value
        End Set
    End Property
    
    Property pbSearchCompanyNameOnly() As Boolean
        Get
            Dim o As Object = ViewState("SearchCompanyNameOnly")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SearchCompanyNameOnly") = Value
        End Set
    End Property
    
    Property psDefaultDescription() As String
        Get
            Dim o As Object = ViewState("DefaultDescription")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("DefaultDescription") = Value
        End Set
    End Property
    
    Property pbMakeRef1Mandatory() As Boolean
        Get
            Dim o As Object = ViewState("MakeRef1Mandatory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("MakeRef1Mandatory") = Value
        End Set
    End Property
    
    Property pbMakeRef2Mandatory() As Boolean
        Get
            Dim o As Object = ViewState("MakeRef2Mandatory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("MakeRef2Mandatory") = Value
        End Set
    End Property
    
    Property pbMakeRef3Mandatory() As Boolean
        Get
            Dim o As Object = ViewState("MakeRef3Mandatory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("MakeRef3Mandatory") = Value
        End Set
    End Property
    
    Property pbMakeRef4Mandatory() As Boolean
        Get
            Dim o As Object = ViewState("MakeRef4Mandatory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("MakeRef4Mandatory") = Value
        End Set
    End Property
    
    Property plThirdPartyCollectionKey() As Long
        Get
            Dim o As Object = ViewState("ThirdPartyCollectionKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("ThirdPartyCollectionKey") = Value
        End Set
    End Property
    
    Property pbHideCollectionButton() As Boolean
        Get
            Dim o As Object = ViewState("HideCollectionButton")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("HideCollectionButton") = Value
        End Set
    End Property
    
    Protected Sub check_DocumentsOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call CheckForNonDocNonEUConsignment()
    End Sub

    Protected Sub drop_CneeCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call CheckForNonDocNonEUConsignment()
    End Sub
    
    Protected Function CheckForNonDocNonEUConsignment()
        If check_DocumentsOnly.Checked = False AndAlso IsNotEUCountry() Then
            btnSubmitNewConsignment.OnClientClick = "return confirm('This is a non-document consignment addressed to a country outside the EU.\n\n Please ensure you request a commercial invoice for Customs purposes in the Special Instructions field.\n\nClick OK to submit the consignment, Cancel to continue editing.');"
        Else
            btnSubmitNewConsignment.OnClientClick = String.Empty
        End If
    End Function
    
    Protected Function IsNotEUCountry() As Boolean
        'AUSTRIA  14, BELGIUM  21, BULGARIA  33, CYPRUS  55, CZECH REPUBLIC  56, DENMARK  57, ESTONIA  67, FINLAND  72, FRANCE  73, GERMANY  81, GREECE  84, HUNGARY  97, IRELAND  103, ITALY  105, LATVIA  117, LITHUANIA  123, LUXEMBOURG  124, ALTA  132, NETHERLANDS  150, POLAND  170, PORTUGAL  171, ROMANIA  175, SLOVAKIA   189, SLOVENIA  190, SPAIN  195, SWEDEN  203, U.K.  222
        Dim arrEUCountryCodes() As Integer = {0, 14, 21, 33, 55, 56, 57, 67, 72, 73, 81, 84, 97, 103, 105, 117, 123, 124, 132, 150, 170, 171, 175, 189, 190, 195, 203, 222}
        IsNotEUCountry = Array.IndexOf(arrEUCountryCodes, drop_CneeCountry.SelectedValue) = -1
    End Function
    

    Protected Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtDescription.Text <> "Documents" AndAlso check_DocumentsOnly.Checked Then
            
        End If
    End Sub
    
</script>
<html xmlns=" http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Courier Collection</title>
    <link href="~/css/sprint.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmCourierCollection" runat="server">
        <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_couriercollection">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <asp:Panel id="pnlCourierBookingStatus" runat="server" visible="False" Width="100%">
                <table style="width:100%; font-family:Verdana; font-size:x-small">
                    <tr>
                        <td style="white-space:nowrap; width:50%" align="left">
                            <asp:Button ID="btnCreateConsignment" runat="server" Tooltip="create a consignment" Text="create a consignment" OnClick="btnCreateConsignment_Click" />
                            &nbsp;&nbsp;&nbsp;
                            <asp:Button ID="btnBookACollection" runat="server" Tooltip="book a collection" Text="book a collection" OnClick="btnBookACollection_Click" />
                            &nbsp;&nbsp;&nbsp;
                            <asp:Button ID="btnRequestSupplies_Click" runat="server" Tooltip="request supplies" Text="request supplies" OnClick="btnRequestSupplies_Click_Click" />
                        </td>
                        <td align="right" style="white-space:nowrap;width:50%">
                            &nbsp;</td>
                    </tr>
                </table>
                <asp:Panel id="pnlCollectionStatus" runat="server" visible="False" Width="100%">
                    <br />&nbsp;
                    <asp:Label id="Label1" runat="server" font-bold="True" font-names="Verdana" font-size="X-Small">You have</asp:Label>
                    <asp:Label id="lblCollectionCount" runat="server" font-bold="True" forecolor="Red" font-names="Verdana" font-size="X-Small"></asp:Label>
                    <asp:Label ID="lblLegendCollection" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small">collection</asp:Label>
                    <asp:Label id="Label8" runat="server" font-bold="True" font-names="Verdana" font-size="X-Small"> booked for today</asp:Label><br />
                    <br />&nbsp;
                    <asp:Label id="lblCollectionMessage" runat="server" font-names="Verdana" font-size="XX-Small"></asp:Label>
                    <br />
                    <asp:DataGrid id="dgTodaysCollections" runat="server" Width="800px" Font-Size="XX-Small" Font-Names="Verdana" OnPageIndexChanged="dgTodaysCollections_Page_Change" AllowPaging="True" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="grd_TodaysCollections_item_click" PageSize="5">
                        <FooterStyle wrap="False"></FooterStyle>
                        <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                        <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                        <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                        <ItemStyle backcolor="LightGray"></ItemStyle>
                        <Columns>
                            <asp:TemplateColumn>
                                <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btnCancelCollection" CommandName="cancel" runat="server" Tooltip="cancel this collection" Text="cancel" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="Key" HeaderText="Number">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                <ItemStyle wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ReadyAt" HeaderText="Pickup Time" DataFormatString="{0:dd/MM HH:mm}">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Company" HeaderText="Company">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                <ItemStyle wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CollectionPoint" HeaderText="Collection Point">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                <ItemStyle wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ContactName" HeaderText="Contact Name">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                <ItemStyle wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="StateId" HeaderText="Status">
                                <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                <ItemStyle wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn>
                                <ItemStyle horizontalalign="Right"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btnDriversManifest" Tooltip="use your browser to print a driver's manifest" OnClientClick='<%#"Javascript:PrintManifest(" & sDriversManifestArgs(Container.DataItem) & ")"%>' runat="server" Text="print driver's manifest" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </asp:Panel>
                <br />&nbsp;
                <asp:Label id="Label7" runat="server" font-bold="True" font-names="Verdana" font-size="X-Small">You
                have</asp:Label>&nbsp;<asp:Label id="lblConsignmentCount" runat="server" font-bold="True" forecolor="Red" font-names="Verdana" font-size="X-Small"></asp:Label>&nbsp;<asp:Label ID="lblLegendConsignment" runat="server" Font-Bold="True" Font-Names="Verdana"
                    Font-Size="X-Small">consignment</asp:Label>
                <asp:Label id="Label2" runat="server" font-bold="True" font-names="Verdana" font-size="X-Small"> awaiting collection</asp:Label>
                <br />
                <br />
                <asp:DataGrid id="dgConsignmentsAwaitingCollection" runat="server" Width="100%" Font-Size="XX-Small" Font-Names="Verdana" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="grd_ConsAwaitingCollection_item_click">
                    <FooterStyle wrap="False"></FooterStyle>
                    <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                    <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                    <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                    <ItemStyle backcolor="LightGray"></ItemStyle>
                    <Columns>
                        <asp:TemplateColumn>
                            <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                            <ItemTemplate>
                                <asp:Button ID="btnEditConsignment" CommandName="edit" runat="server" Tooltip="edit this consignment" Text="edit" />
                                <asp:Button ID="btnDeleteConsignment" CommandName="delete" runat="server" Tooltip="delete this consignment" Text="delete" />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:BoundColumn DataField="Key" SortExpression="Key" HeaderText="Consignment">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemStyle wrap="False"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn HeaderText="Collection">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemTemplate>
                                <asp:LinkButton ID="LinkButton1" runat="server" CommandName="collection" ForeColor="Blue">
                                    <%# DataBinder.Eval(Container.DataItem,"CourierBookingKey") %>
                                </asp:LinkButton>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:BoundColumn DataField="CustomerRef1" SortExpression="CustomerRef1" HeaderText="Shipper's Ref">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemStyle wrap="False"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="CneeName" SortExpression="CneeName" HeaderText="Addressed To">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemStyle wrap="False"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="City">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemStyle wrap="False"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                            <ItemStyle wrap="False"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="SpecialInstructions" SortExpression="SpecialInstructions" HeaderText="Special Instructions">
                            <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn>
                            <ItemStyle horizontalalign="Right"></ItemStyle>
                            <ItemTemplate>
                                <asp:Button ID="btnPrintAWB" OnClientClick='<%# "Javascript:PrintConsignment(" & DataBinder.Eval(Container.DataItem,"Key") & ")" %>' runat="server" Text="print awb" />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn>
                            <ItemStyle horizontalalign="Right"></ItemStyle>
                            <ItemTemplate>
                                    <asp:Button ID="btnPrintLabel" OnClientClick='<%# "Javascript:PrintLabel(" & DataBinder.Eval(Container.DataItem ,"Key") & ")" %> ' runat="server" Text="print label" />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
                &nbsp;
                <asp:LinkButton id="LinkButton1" onclick="btn_RefreshStatusPage_click" runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Blue" CausesValidation="False">refresh</asp:LinkButton>
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            </asp:Panel>
            <asp:Panel id="pnlSearchAddressBook" runat="server" visible="False" Width="100%">
            <asp:Label ID="Label33z" runat="server" font-size="XX-Small" font-names="Verdana" Font-Bold="true">Search Address Book</asp:Label>
                <asp:Table id="tabSearchAddressBook" runat="server" Width="100%" Font-Size="XX-Small" Font-Names="Verdana">
                    <asp:TableRow>
                        <asp:TableCell Width="20px"></asp:TableCell>
                        <asp:TableCell Width="580px"></asp:TableCell>
                        <asp:TableCell Width="100px">
                            <asp:Button ID="btnGoBackToConsignment" runat="server" Tooltip="go back" Text="go back" OnClick="btnGoBackToConsignment_Click" CausesValidation="false" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="20px"></asp:TableCell>
                        <asp:TableCell Width="580px"></asp:TableCell>
                        <asp:TableCell Width="100px">
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label3" runat="server" font-size="XX-Small" font-names="Verdana">Search my address book: </asp:Label>
                            <asp:TextBox runat="server" Height="20px" Width="150px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtSearchCriteriaAddress" Tooltip="search all my addresses"></asp:TextBox>
                            <asp:Button ID="btnGo" runat="server" Tooltip="search" Text="go" OnClick="btnGo_Click" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnShowAllAddresses" runat="server" Tooltip="show all addresses" Text="show all" OnClick="btnShowAllAddresses_Click" />
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right">
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell ColumnSpan="2">
                            <br/>
                            <asp:DataGrid id="dgAddressBook" runat="server" Width="100%" Font-Size="XX-Small" Font-Names="Verdana" OnPageIndexChanged="dgAddressBook_Page_Change" AllowPaging="True" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="dgAddressBook_item_click" PageSize="12">
                                <FooterStyle wrap="False"></FooterStyle>
                                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                <ItemStyle backcolor="LightGray"></ItemStyle>
                                <Columns>
                                    <asp:TemplateColumn>
                                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                                        <ItemTemplate>
                                            <asp:Button ID="btnSelectAddress" CommandName="select" runat="server" Tooltip="select this address" Text="select" />
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:BoundColumn Visible="False" DataField="DestKey">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Company" HeaderText="Company">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Addr1" HeaderText="Addr1">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Town" HeaderText="City">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CountryName" HeaderText="CountryName">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell ColumnSpan="2">
                            <asp:Label id="lblAddressMessage" runat="server" forecolor="Red" font-names="Verdana" font-size="X-Small"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <br />
            </asp:Panel>
            <asp:Panel id="pnlCreateConsignment" runat="server" visible="false" Width="100%">
                <table id="tabContainer" runat="server" style="width:700px">
                    <tr>
                        <td style="width:20px"></td>
                        <td style="width:680px">
                            <table id="tabHeader" runat="server" style=" font-size:x-Small; font-family:Verdana">
                                <tr>
                                    <td style="width:140px"></td>
                                    <td style="width:190px"></td>
                                    <td style="width:20px"></td>
                                    <td style="width:140px"></td>
                                    <td style="width:190px"></td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <asp:Label ID="Label73" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Create a Consignment</asp:Label>
                                    </td>
                                    <td align="Right" colspan="3">
                                        <asp:Button ID="btnShowCourierBookingStatus0" runat="server" Tooltip="back to main status page" Text="go back" OnClick="btnShowCourierBookingStatus_Click" CausesValidation="false" />
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label4" runat="server" font-size="X-Small" font-names="Verdana">Delivery Address</asp:Label>
                                    </td>
                                    <td>
                                        <asp:LinkButton ID="LinkButton2" onclick="btn_GetFromAddressBook_click" TabIndex="-1" runat="server" Causesvalidation="False" forecolor="blue" font-size="XX-Small" font-names="Verdana">get from my address book</asp:LinkButton>
                                    </td>
                                    <td align="Right" colspan="3">
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="Label5" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Company:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtCneeName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="1" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeName" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="Label6" runat="server" font-size="XX-Small" font-names="Verdana">Post Code:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="6" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneePostCode" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="height: 22px">
                                        <asp:Label ID="Label9" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Addr 1:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtCneeAddr1" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="2" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeAddr1" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                    <td style="height: 22px"></td>
                                    <td style="height: 22px">
                                        <asp:Label ID="lbl0005" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Country:</asp:Label> &nbsp;
                                        <asp:CompareValidator ID="CompareValidator1" runat="server" ValueToCompare="0" Operator="NotEqual" Font-Names="Verdana" ControlToValidate="drop_CneeCountry" Font-Size="XX-Small">#</asp:CompareValidator>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:DropDownList runat="server" ID="drop_CneeCountry" TabIndex="7" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" style="width:190px" AutoPostBack="True" OnSelectedIndexChanged="drop_CneeCountry_SelectedIndexChanged" Width="100%"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td style="height: 22px">
                                        <asp:Label ID="lbl0006" runat="server" font-size="XX-Small" font-names="Verdana">Addr 2:</asp:Label>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="3" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeAddr2" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                    <td style="height: 22px"></td>
                                    <td style="height: 22px">
                                        <asp:Label ID="lbl0007" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Attn
                                        of:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="txtCneeCTCName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="8" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeCTCName" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="lbl0009" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">City:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="txtCneeCity" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="4" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeCity" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0008" runat="server" font-size="XX-Small" font-names="Verdana">Contact Tel:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="9" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeTel" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="lbl0010" runat="server" font-size="XX-Small" font-names="Verdana">State/County:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeState" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0011" runat="server" font-size="XX-Small" font-names="Verdana">Contact Email:</asp:Label>
                                        <asp:RegularExpressionValidator ID="revCneeEmail" runat="server" ErrorMessage="#" ControlToValidate="txtCneeEmail"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="10" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCneeEmail" MaxLength="50" style="width:190px" Width="100%"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="Top">
                                        <asp:Label ID="lbl0012" runat="server" font-size="XX-Small" font-names="Verdana">Insurance</asp:Label>
                                    </td>
                                    <td rowspan="2">
                                        <asp:Label ID="lbl0013" runat="server" font-size="XX-Small" font-names="Verdana" width="190px">Please note that our
                                        liability for loss or damage is limited to £50 unless a value has been entered in
                                        the 'Value for Insurance' box.</asp:Label>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0014" runat="server" font-size="XX-Small" font-names="Verdana">Val for Insurance:</asp:Label>
                                        &nbsp;
                                        <asp:CustomValidator
                                            ID="cvValForInsurance" ValidationGroup="vgCreateConsignment" ControlToValidate="txtValForInsurance" EnableClientScript="false" OnServerValidate="CheckValidNumber_ServerValidate" runat="server" ErrorMessage="#" Text="#"/>
                                    </td>
                                    <td valign="Top">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="11" Font-Size="XX-Small" Font-Names="Verdana" ID="txtValForInsurance" width="50px"></asp:TextBox>&nbsp;
                                        <asp:Label ID="lbl0015" runat="server" font-size="XX-Small" font-names="Verdana">Currency</asp:Label>&nbsp;
                                         <asp:DropDownList runat="server" id="drop_ValForInsurance" TabIndex="-1" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana">
                                            <asp:ListItem Value="123"> £ </asp:ListItem>
                                            <asp:ListItem Value="52"> € </asp:ListItem>
                                            <asp:ListItem Value="168"> $ </asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td></td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0016" runat="server" font-size="XX-Small" font-names="Verdana">Val for Customs:</asp:Label>
                                        &nbsp;
                                        <asp:CustomValidator ID="cvValForCustoms" ValidationGroup="vgCreateConsignment" ControlToValidate="txtValForCustoms" EnableClientScript="false" OnServerValidate="CheckValidNumber_ServerValidate" runat="server" ErrorMessage="#" Text="#"/>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana" ID="txtValForCustoms" width="50px"></asp:TextBox>&nbsp;
                                        <asp:Label ID="lbl0017" runat="server" font-size="XX-Small" font-names="Verdana">Currency</asp:Label>&nbsp;
                                         <asp:DropDownList runat="server" id="drop_ValForCustoms" TabIndex="-1" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana">
                                            <asp:ListItem Value="123"> £ </asp:ListItem>
                                            <asp:ListItem Value="52"> € </asp:ListItem>
                                            <asp:ListItem Value="168"> $ </asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label runat="server" id="lblRef1" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_Ref1" runat="server" ControlToValidate="txtCustRef1" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCustRef1" MaxLength="30" width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0019" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">No of Pieces:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="txtNoPieces" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                        <asp:RangeValidator ID="rvTxtNoPieces" runat="server" Type="Integer" MinimumValue="1" MaximumValue="9999" ControlToValidate="txtNoPieces">#</asp:RangeValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana" id="txtNoPieces" width="30px">1</asp:TextBox>
                                        <asp:Label ID="lbl0021" runat="server" font-size="XX-Small" font-names="Verdana">Weight (Kgs):</asp:Label>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="14" Font-Size="XX-Small" Font-Names="Verdana" id="txtWeight" width="60px"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label runat="server" id="lblRef2" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_Ref2" runat="server" ControlToValidate="txtCustRef2" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="17" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCustRef2" MaxLength="30" width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        <asp:Label ID="lbl0023" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Description:</asp:Label> 
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="txtDescription" Font-Size="XX-Small">#</asp:RequiredFieldValidator></td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="15" Font-Size="XX-Small" Font-Names="Verdana" id="txtDescription" width="100%" OnTextChanged="txtDescription_TextChanged"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label runat="server" id="lblRef3" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_Ref3" runat="server" ControlToValidate="txtCustRef3" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="18" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCustRef3" MaxLength="50" width="100%"/>
                                    </td>
                                    <td></td>
                                    <td>
                                        &nbsp; &nbsp;
                                    </td>
                                    <td>
                                        <asp:CheckBox id="check_DocumentsOnly" TabIndex="-1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Docs only" Checked="True" AutoPostBack="True" OnCheckedChanged="check_DocumentsOnly_CheckedChanged"></asp:CheckBox></td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label runat="server" id="lblRef4" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_Ref4" runat="server" ControlToValidate="txtCustRef4" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="19" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCustRef4" MaxLength="50" width="100%"/>
                                    </td>
                                    <td></td>
                                    <td align="Right" colspan="2">
                                        <asp:Label ID="lbl0025" runat="server" font-size="XX-Small" font-names="Verdana">see our</asp:Label>
                                        &nbsp;<asp:HyperLink ID="HyperLink4" runat="server" NavigateUrl="ConditionsOfCarriage.pdf" TabIndex="-1" ForeColor="Blue" Font-Size="XX-Small" Font-Names="Verdana" Target="_blank">conditions of carriage</asp:HyperLink>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:Label ID="lbl0026" runat="server" font-size="XX-Small" font-names="Verdana">Special Instructions:</asp:Label>
                                    </td>
                                    <td colspan="4">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="20" Font-Size="XX-Small" Font-Names="Verdana" id="txtSpecialInstructions" width="100%"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td></td>
                                    <td colspan="4">
                                        <asp:Label ID="lbl0027" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Please confirm any special intructions with our customer services department - thank you.</asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                    <td></td>
                                </tr>
                            </table>
                            <table id="tabFooter" runat="server" style="width:680px; font-size:x-Small; font-family:Verdana">
                                <tr>
                                    <td style="width:120px"></td>
                                    <td style="width:30px"></td>
                                    <td style="width:100px"></td>
                                    <td style="width:190px"></td>
                                    <td style="width:240px"></td>
                                </tr>
                                <tr id="Tr1" runat="server" visible="false">
                                    <td>
                                        <asp:Label ID="Label10" runat="server" font-size="XX-Small" font-names="Verdana">SMS Tracking:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:CheckBox Enabled="False" id="chk_SMSTextTracking" TabIndex="-1" runat="server"></asp:CheckBox>
                                    </td>
                                    <td>
                                        <asp:Label ID="lbl0028" runat="server" font-size="XX-Small" font-names="Verdana">Telephone No:</asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox Enabled="False" runat="server" ForeColor="Navy" TabIndex="-1" Font-Size="XX-Small" Font-Names="Verdana" ID="txtTrackingMobileNo" width="190px">coming soon</asp:TextBox>
                                    </td>
                                    <td align="Right"></td>
                                </tr>
                                <tr id="Tr2" runat="server" visible="false">
                                    <td style="height: 34px">
                                        <asp:Label ID="lbl0029" runat="server" font-size="XX-Small" font-names="Verdana">Email Tracking:</asp:Label>
                                    </td>
                                    <td style="height: 34px">
                                        <asp:CheckBox Enabled="False" id="chk_EmailTracking" TabIndex="-1" runat="server"></asp:CheckBox>
                                    </td>
                                    <td style="height: 34px">
                                        <asp:Label ID="lbl0030" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td style="height: 34px">
                                        <asp:TextBox Enabled="False" runat="server" ForeColor="Navy" TabIndex="-1" Font-Size="XX-Small" Font-Names="Verdana" ID="txtTrackingEmailAddr" width="190px">coming soon</asp:TextBox>
                                    </td>
                                    <td align="Right" style="height: 34px">
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                    </td>
                                    <td>
                                    </td>
                                    <td>
                                    </td>
                                    <td>
                                    </td>
                                    <td align="Right">
                                        <asp:Label ID="lbl0031" runat="server" font-size="XX-Small" font-names="Verdana">Save address</asp:Label>
                                        <asp:CheckBox id="chk_SaveAddress" runat="server" TabIndex="-1" Font-Names="Verdana" Font-Size="XX-Small" Checked="False"></asp:CheckBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert01" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td>
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation01" runat="server" Text="Confirm Receipt"
                                            ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td>
                                        <asp:Label ID="lbl0033a" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr01" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert01" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr01"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr id="trPreAlert02" runat="server" visible="false">
                                    <td>
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert02" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td>
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation02" runat="server" Text="Confirm Receipt" ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td>
                                        <asp:Label ID="lbl0033b" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr02" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert02" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr02"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr id="trPreAlert03" runat="server" visible="false">
                                    <td style="height: 22px">
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert03" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td style="height: 22px">
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation03" runat="server" Text="Confirm Receipt" ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td style="height: 22px">
                                        <asp:Label ID="lbl0033c" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr03" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left" style="height: 22px">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert03" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr03"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr id="trPreAlert04" runat="server" visible="false">
                                    <td style="height: 22px">
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert04" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td style="height: 22px">
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation04" runat="server" Text="Confirm Receipt" ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td style="height: 22px">
                                        <asp:Label ID="lbl0033d" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td style="height: 22px">
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr04" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left" style="height: 22px">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert04" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr04"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr id="trPreAlert05" runat="server" visible="false">
                                    <td>
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert05" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td>
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation05" runat="server" Text="Confirm Receipt" ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td>
                                        <asp:Label ID="lbl0033e" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr05" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert05" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr05"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr id="trPreAlert06" runat="server" visible="false">
                                    <td>
                                        &nbsp;<asp:CheckBox id="cbEmailPreAlert06" runat="server" TabIndex="-1" Text="Pre-alert" Font-Names="Verdana" Font-Size="XX-Small"></asp:CheckBox></td>
                                    <td>
                                        &nbsp;<asp:CheckBox ID="cbEmailDeliveryConfirmation06" runat="server" Text="Confirm Receipt" ToolTip="select this checkbox to send a delivery (receipt) confirmation to the email address specified" Font-Names="Verdana" Font-Size="XX-Small" /></td>
                                    <td>
                                        <asp:Label ID="lbl0033f" runat="server" font-size="XX-Small" font-names="Verdana">Email Addr: </asp:Label>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" TabIndex="-1" Font-Names="Verdana" ID="tbPreAlertEmailAddr06" width="190px"></asp:TextBox>
                                    </td>
                                    <td align="left">
                                        <asp:RegularExpressionValidator ID="revEmailPreAlert06" runat="server" ErrorMessage="#" ControlToValidate="tbPreAlertEmailAddr06"
                                            ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"/>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                    </td>
                                    <td colspan="2">
                                        <asp:LinkButton ID="lnkbtnAddAnotherPreAlert" runat="server" Font-Names="Verdana"
                                            Font-Size="XX-Small" OnClick="lnkbtnAddAnotherPreAlert_Click" CausesValidation="False">add another pre-alert</asp:LinkButton></td>
                                    <td>
                                    </td>
                                    <td align="Right">
                                        <asp:Button ID="btnCancelConsignment" runat="server" TabIndex="22" Tooltip="reset page and return to main status page" Causesvalidation="False" Text="cancel" OnClick="btnCancelConsignment_Click" />
                                        &nbsp;&nbsp;
                                        <asp:Button ID="btnSubmitNewConsignment" runat="server" TabIndex="22" Tooltip="submit this page to create a new consignment" Text="submit" OnClick="btnSubmitNewConsignment_Click" />
                                     </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:Panel id="pnlCreateAdHocCollection" runat="server" visible="False" Width="100%">
                <asp:Table id="tabCreateAdHocCollection" runat="server" Width="750px" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow VerticalAlign="Middle">
                        <asp:TableCell VerticalAlign="Middle" HorizontalAlign="Left" Wrap="False">
                            <asp:Label ID="Label11" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Book
                            a Collection</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell VerticalAlign="Middle" HorizontalAlign="Right" Wrap="False">
                            <asp:Button ID="btnShowCourierBookingStatus2" runat="server" Tooltip="back to main status page" Text="go back" OnClick="btnShowCourierBookingStatus_Click" CausesValidation="false" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow VerticalAlign="Middle">
                        <asp:TableCell VerticalAlign="Middle" ColumnSpan="2" HorizontalAlign="Left">
                            <asp:Label ID="Label12" runat="server" font-size="XX-Small" font-names="Verdana">Use this page
                            to schedule a driver to collect your consignment(s) from the address below. You can
                            change the default address by overtyping the text below. Please take care to amend the time and date your consignment(s) will
                            be available for collection. If you require collection outside of our normal working hours, please confirm your collection by telephoning our Customer Services Department.</asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <asp:Table id="tabAdHocCollection" runat="server" Width="750px" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Wrap="False" Width="180px">
                            <asp:Label ID="Label13" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Company:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="txtCollectionName" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell Width="180px">
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="1" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionName"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell Width="180px">
                            <asp:Label ID="Label14" runat="server" font-size="XX-Small" forecolor="Red" font-names="Verdana">Time
                            Ready:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="180px">
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="9" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_HourReady">
                                <asp:ListItem Value="1">01</asp:ListItem>
                                <asp:ListItem Value="2">02</asp:ListItem>
                                <asp:ListItem Value="3">03</asp:ListItem>
                                <asp:ListItem Value="4">04</asp:ListItem>
                                <asp:ListItem Value="5">05</asp:ListItem>
                                <asp:ListItem Value="6">06</asp:ListItem>
                                <asp:ListItem Value="7">07</asp:ListItem>
                                <asp:ListItem Value="8">08</asp:ListItem>
                                <asp:ListItem Value="9">09</asp:ListItem>
                                <asp:ListItem Value="10">10</asp:ListItem>
                                <asp:ListItem Value="11">11</asp:ListItem>
                                <asp:ListItem Value="12">12</asp:ListItem>
                                <asp:ListItem Value="13">13</asp:ListItem>
                                <asp:ListItem Value="14">14</asp:ListItem>
                                <asp:ListItem Value="15">15</asp:ListItem>
                                <asp:ListItem Value="16">16</asp:ListItem>
                                <asp:ListItem Value="17">17</asp:ListItem>
                                <asp:ListItem Value="18">18</asp:ListItem>
                                <asp:ListItem Value="19">19</asp:ListItem>
                                <asp:ListItem Value="20">20</asp:ListItem>
                                <asp:ListItem Value="21">21</asp:ListItem>
                                <asp:ListItem Value="22">22</asp:ListItem>
                                <asp:ListItem Value="23">23</asp:ListItem>
                                <asp:ListItem Value="24">24</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="10" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_MinuteReady">
                                <asp:ListItem Value="1">00</asp:ListItem>
                                <asp:ListItem Value="2">15</asp:ListItem>
                                <asp:ListItem Value="3">30</asp:ListItem>
                                <asp:ListItem Value="4">45</asp:ListItem>
                            </asp:DropDownList>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label15" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Addr
                            1:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ControlToValidate="txtCollectionAddr1" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="2" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionAddr1"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label16" runat="server" font-size="XX-Small" forecolor="Red" font-names="Verdana">Date
                            Ready:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="11" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_DayReady">
                                <asp:ListItem Value="1">1</asp:ListItem>
                                <asp:ListItem Value="2">2</asp:ListItem>
                                <asp:ListItem Value="3">3</asp:ListItem>
                                <asp:ListItem Value="4">4</asp:ListItem>
                                <asp:ListItem Value="5">5</asp:ListItem>
                                <asp:ListItem Value="6">6</asp:ListItem>
                                <asp:ListItem Value="7">7</asp:ListItem>
                                <asp:ListItem Value="8">8</asp:ListItem>
                                <asp:ListItem Value="9">9</asp:ListItem>
                                <asp:ListItem Value="10">10</asp:ListItem>
                                <asp:ListItem Value="11">11</asp:ListItem>
                                <asp:ListItem Value="12">12</asp:ListItem>
                                <asp:ListItem Value="13">13</asp:ListItem>
                                <asp:ListItem Value="14">14</asp:ListItem>
                                <asp:ListItem Value="15">15</asp:ListItem>
                                <asp:ListItem Value="16">16</asp:ListItem>
                                <asp:ListItem Value="17">17</asp:ListItem>
                                <asp:ListItem Value="18">18</asp:ListItem>
                                <asp:ListItem Value="19">19</asp:ListItem>
                                <asp:ListItem Value="20">20</asp:ListItem>
                                <asp:ListItem Value="21">21</asp:ListItem>
                                <asp:ListItem Value="22">22</asp:ListItem>
                                <asp:ListItem Value="23">23</asp:ListItem>
                                <asp:ListItem Value="24">24</asp:ListItem>
                                <asp:ListItem Value="25">25</asp:ListItem>
                                <asp:ListItem Value="26">26</asp:ListItem>
                                <asp:ListItem Value="27">27</asp:ListItem>
                                <asp:ListItem Value="28">28</asp:ListItem>
                                <asp:ListItem Value="29">29</asp:ListItem>
                                <asp:ListItem Value="30">30</asp:ListItem>
                                <asp:ListItem Value="31">31</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_MonthReady">
                                <asp:ListItem Value="1">JAN</asp:ListItem>
                                <asp:ListItem Value="2">FEB</asp:ListItem>
                                <asp:ListItem Value="3">MAR</asp:ListItem>
                                <asp:ListItem Value="4">APR</asp:ListItem>
                                <asp:ListItem Value="5">MAY</asp:ListItem>
                                <asp:ListItem Value="6">JUN</asp:ListItem>
                                <asp:ListItem Value="7">JUL</asp:ListItem>
                                <asp:ListItem Value="8">AUG</asp:ListItem>
                                <asp:ListItem Value="9">SEP</asp:ListItem>
                                <asp:ListItem Value="10">OCT</asp:ListItem>
                                <asp:ListItem Value="11">NOV</asp:ListItem>
                                <asp:ListItem Value="12">DEC</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_YearReady"></asp:DropDownList>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label17" runat="server" font-size="XX-Small" font-names="Verdana">Addr 2:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="3" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionAddr2"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label18" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Contact
                            Name:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" ControlToValidate="txtCollectionCTCName" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="14" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionCTCName"></asp:TextBox>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label19" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">City:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ControlToValidate="txtCollectionCity" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="4" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionCity"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label20" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Contact
                            Tel:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="txtCollectionTel" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="15" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionTel"></asp:TextBox>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label21" runat="server" font-size="XX-Small" font-names="Verdana">County:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="5" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionState"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label22" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Collection Point:</asp:Label> &nbsp;
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" ControlToValidate="txtCollectionPoint" Font-Size="XX-Small">###</asp:RequiredFieldValidator>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="16" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionPoint"></asp:TextBox>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label23" runat="server" font-size="XX-Small" font-names="Verdana">Post Code:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="6" Width="180px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtCollectionPostCode"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label24" runat="server" font-size="XX-Small" font-names="Verdana">Country:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell>
                            <asp:DropDownList runat="server" ForeColor="Navy" TabIndex="7" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_CollectionCountry"></asp:DropDownList>
                        </asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell></asp:TableCell>
                        <asp:TableCell>
                            <asp:Label ID="Label25" runat="server" font-size="XX-Small" font-names="Verdana">Note to Driver:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell ColumnSpan="2">
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="8" Width="350px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtNoteToDriver"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right">
                            <asp:Button ID="btnCancelCollection" runat="server" Tooltip="cancel Collection" CausesValidation="False" Text="cancel" OnClick="btnCancelCollection_Click" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnSubmitNewCollection" runat="server" Tooltip="submit this page to create a new Collection" Text="submit" OnClick="btnSubmitNewCollection_Click" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="5">
                            <br />
                            <asp:Label ID="Label26" runat="server" font-size="XX-Small" font-names="Verdana">N.B. If you wish
                            to record special instructions for your consignment(s) then please do so when creating
                            the consignment record - these are collection instructions only.</asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlEditConsignment" runat="server" visible="False" Width="100%">
                <asp:Table id="tabEditContainer" runat="server" Width="700px">
                    <asp:TableRow>
                        <asp:TableCell Width="20px"></asp:TableCell>
                        <asp:TableCell Width="680px">
                            <asp:Table id="tabEditHeader" runat="server" Font-Size="X-Small" Font-Names="Verdana">
                                <asp:TableRow>
                                    <asp:TableCell Width="480px">
                                        <asp:Label ID="Label27" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Editing Consignment No</asp:Label> &nbsp;<asp:Label runat="server" id="lblConNoteNumber" font-size="X-Small" forecolor="Red" font-names="Verdana" font-bold="True"></asp:Label>
                                    </asp:TableCell>
                                    <asp:TableCell HorizontalAlign="Right" Width="200px">
                                        <asp:Button ID="btnShowCourierBookingStatus3" runat="server" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" CausesValidation="false" />
                                    </asp:TableCell>
                                </asp:TableRow>
                            </asp:Table>
                            <asp:Table id="tabEditConsignment" runat="server" Font-Size="XX-Small" Font-Names="Verdana">
                                <asp:TableRow>
                                    <asp:TableCell Width="140px"></asp:TableCell>
                                    <asp:TableCell Width="190px"></asp:TableCell>
                                    <asp:TableCell Width="20px"></asp:TableCell>
                                    <asp:TableCell Width="140px"></asp:TableCell>
                                    <asp:TableCell Width="190px"></asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label28" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">From:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ControlToValidate="txtEditCnorName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorName" TabIndex="1" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label29" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">To:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" ControlToValidate="txtEditCneeName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeName" TabIndex="10" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label30" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Addr 1:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator15" runat="server" ControlToValidate="txtEditCnorAddr1" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorAddr1" TabIndex="2" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label31" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Addr 1:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator16" runat="server" ControlToValidate="txtEditCneeAddr1" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeAddr1" TabIndex="11" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label32" runat="server" font-size="XX-Small" font-names="Verdana">Addr 2:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorAddr2" TabIndex="3" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label33" runat="server" font-size="XX-Small" font-names="Verdana">Addr 2:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeAddr2" TabIndex="12" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label34" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">City:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator17" runat="server" ControlToValidate="txtEditCnorCity" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorCity" TabIndex="4" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label35" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">City:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator18" runat="server" ControlToValidate="txtEditCneeCity" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeCity" TabIndex="13" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label36" runat="server" font-size="XX-Small" font-names="Verdana">State:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorState" TabIndex="5" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label37" runat="server" font-size="XX-Small" font-names="Verdana">State:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeState" TabIndex="14" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label38" runat="server" font-size="XX-Small" font-names="Verdana">Post Code:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorPostCode" TabIndex="6" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label39" runat="server" font-size="XX-Small" font-names="Verdana">Post Code:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneePostCode" TabIndex="15" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label40" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Country:</asp:Label>&nbsp;
                                        &nbsp;<asp:CompareValidator ID="CompareValidator2" runat="server" ValueToCompare="0" Operator="NotEqual" Font-Names="Verdana" ControlToValidate="drop_EditCnorCountry" Font-Size="XX-Small">#</asp:CompareValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:DropDownList runat="server" ForeColor="Navy" Width="190px" TabIndex="7" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_EditCnorCountry"></asp:DropDownList>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label41" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Country:</asp:Label>&nbsp;
                                        &nbsp;<asp:CompareValidator ID="CompareValidator3" runat="server" ValueToCompare="0" Operator="NotEqual" Font-Names="Verdana" ControlToValidate="drop_EditCneeCountry" Font-Size="XX-Small">#</asp:CompareValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:DropDownList runat="server" ForeColor="Navy" Width="190px" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana" ID="drop_EditCneeCountry"></asp:DropDownList>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label42" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Ctc Name:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator19" runat="server" ControlToValidate="txtEditCnorCtCName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorCtCName" TabIndex="8" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label43" runat="server" font-size="XX-Small" font-names="Verdana" forecolor="Red">Attn:</asp:Label>&nbsp;
                                        &nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator20" runat="server" ControlToValidate="txtEditCneeCtCName" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeCtCName" TabIndex="17" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label44" runat="server" font-size="XX-Small" font-names="Verdana">Tel:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCnorTel" TabIndex="9" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label45" runat="server" font-size="XX-Small" font-names="Verdana">Tel:</asp:Label>&nbsp;
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" font-size="XX-Small" font-names="Verdana" id="txtEditCneeTel" TabIndex="18" width="190px" MaxLength="50" ForeColor="Navy"/>
                                    </asp:TableCell>
                                </asp:TableRow>

                                <asp:TableRow>
                                    <asp:TableCell VerticalAlign="Top">
                                        <asp:Label ID="Label46" runat="server" font-size="XX-Small" font-names="Verdana">Insurance</asp:Label>
                                    </asp:TableCell>
                                    <asp:TableCell RowSpan="2">
                                        <asp:Label ID="Label47" runat="server" font-size="XX-Small" font-names="Verdana" width="190px">Please note that our
                                        liability for loss or damage is limited to £50 unless a value has been entered in
                                        the 'Value for Insurance' box.</asp:Label>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label48" runat="server" font-size="XX-Small" font-names="Verdana">Val for Insurance:</asp:Label>
                                        &nbsp;
                                        <asp:CustomValidator
                                            ID="cvEditValueForInsurance" ValidationGroup="vgEditConsignment" ControlToValidate="txtEditValueForInsurance" EnableClientScript="false" OnServerValidate="CheckValidNumber_ServerValidate" runat="server" ErrorMessage="#" Text="#"/>
                                    </asp:TableCell>
                                    <asp:TableCell VerticalAlign="Top">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="19" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditValueForInsurance" MaxLength="12" width="50px"/>
                                        &nbsp;
                                        <asp:Label ID="Label49" runat="server" font-size="XX-Small" font-names="Verdana">Currency</asp:Label>&nbsp;
                                         <asp:DropDownList runat="server" id="drop_EditValForInsurance" TabIndex="-1" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana">
                                            <asp:ListItem Value="123"> £ </asp:ListItem>
                                            <asp:ListItem Value="52"> € </asp:ListItem>
                                            <asp:ListItem Value="168"> $ </asp:ListItem>
                                        </asp:DropDownList>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label50" runat="server" font-size="XX-Small" font-names="Verdana">Val for Customs:</asp:Label>
                                        &nbsp;
                                        <asp:CustomValidator
                                            ID="cvEditValueForCustoms" ValidationGroup="vgEditConsignment" ControlToValidate="txtEditValueForCustoms" EnableClientScript="false" OnServerValidate="CheckValidNumber_ServerValidate" runat="server" ErrorMessage="#" Text="#"/>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="20" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditValueForCustoms" MaxLength="12" width="50px"/>
                                        &nbsp;
                                        <asp:Label ID="Label51" runat="server" font-size="XX-Small" font-names="Verdana">Currency</asp:Label>&nbsp;
                                         <asp:DropDownList runat="server" id="drop_EditValForCustoms" TabIndex="-1" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana">
                                            <asp:ListItem Value="123"> £ </asp:ListItem>
                                            <asp:ListItem Value="52"> € </asp:ListItem>
                                            <asp:ListItem Value="168"> $ </asp:ListItem>
                                        </asp:DropDownList>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label runat="server" id="lblEditRef1" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_EditRef1" runat="server" ControlToValidate="txtEditCustRef1" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="24" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditCustRef1" MaxLength="30" width="190px"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label52" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">No
                                        of Pieces:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator21" runat="server" ControlToValidate="txtEditNumPieces" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                        <asp:RangeValidator ID="rvTxtEditNumPieces" runat="server" Type="Integer" MinimumValue="1" MaximumValue="9999" ControlToValidate="txtEditNumPieces">#</asp:RangeValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="21" Font-Size="XX-Small" Font-Names="Verdana" id="txtEditNumPieces" MaxLength="6" width="30px">1</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label runat="server" id="lblEditRef2" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_EditRef2" runat="server" ControlToValidate="txtEditCustRef2" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="25" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditCustRef2" MaxLength="30" width="190px"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label53" runat="server" font-size="XX-Small" font-names="Verdana">Weight (Kgs):</asp:Label>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="22" Font-Size="XX-Small" Font-Names="Verdana" id="txtEditWeight" width="80px"></asp:TextBox>
                                        &nbsp;<asp:CheckBox id="check_EditDocumentsOnly" TabIndex="-1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Docs only" Checked="True"></asp:CheckBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label runat="server" id="lblEditRef3" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_EditRef3" runat="server" ControlToValidate="txtEditCustRef3" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="26" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditCustRef3" MaxLength="50" width="190px"></asp:TextBox>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell>
                                        <asp:Label ID="Label54" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Description:</asp:Label> &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator22" runat="server" ControlToValidate="txtEditDescription" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="23" Font-Size="XX-Small" Font-Names="Verdana" id="txtEditDescription" MaxLength="250" width="190px">Documents</asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label runat="server" id="lblEditRef4" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>&nbsp;
                                        <asp:RequiredFieldValidator id="validator_EditRef4" runat="server" ControlToValidate="txtEditCustRef4" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                                    </asp:TableCell>
                                    <asp:TableCell>
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="27" Font-Size="XX-Small" Font-Names="Verdana" ID="txtEditCustRef4" MaxLength="50" width="190px"/>
                                    </asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell HorizontalAlign="Right" ColumnSpan="2">
                                        <asp:Label ID="Label55" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana">see our</asp:Label>
                                        &nbsp;<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="conditions_of_carriage.html" TabIndex="-1" ForeColor="Blue" Font-Size="XX-Small" Font-Names="Verdana" Target="_blank">conditions of carriage</asp:HyperLink>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell>
                                        <asp:Label ID="Label56" runat="server" font-size="XX-Small" font-names="Verdana">Special Instructions:</asp:Label>
                                    </asp:TableCell>
                                    <asp:TableCell ColumnSpan="4">
                                        <asp:TextBox runat="server" ForeColor="Navy" TabIndex="28" Font-Size="XX-Small" Font-Names="Verdana" id="txtEditSpclInstructions" MaxLength="950" width="550px"></asp:TextBox>
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell ColumnSpan="4">
                                        <asp:Label ID="Label57" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana">Please confirm any special intructions with our customer services department - thank you.</asp:Label>
                                    </asp:TableCell>
                                </asp:TableRow>
                            </asp:Table>

                            <asp:Table id="tabEditFooter" runat="server" Width="680px" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                                <asp:TableRow>
                                    <asp:TableCell width="120px"></asp:TableCell>
                                    <asp:TableCell width="30px"></asp:TableCell>
                                    <asp:TableCell width="100px"></asp:TableCell>
                                    <asp:TableCell width="190px"></asp:TableCell>
                                    <asp:TableCell width="240px"></asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell></asp:TableCell>
                                    <asp:TableCell HorizontalAlign="Right">
                                        <asp:Button ID="btnShowCourierBookingStatus4" runat="server" TabIndex="29" CausesValidation="False" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="cancel" />
                                        &nbsp;&nbsp;
                                        <asp:Button ID="btnSaveConsignmentChanges" runat="server" TabIndex="30" Tooltip="save consignment changes" Text="save" OnClick="btnSaveConsignmentChanges_Click" />
                                    </asp:TableCell>
                                </asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell ColumnSpan="5">
                                        <asp:Label id="lblEditDateError" runat="server" font-size="X-Small" font-names="Verdana" forecolor="#00C000"></asp:Label>
                                    </asp:TableCell>
                                </asp:TableRow>
                            </asp:Table>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlRequestSupplies" runat="server" visible="False" Width="100%">
                <asp:Table id="tabRequestSupplies" runat="server" Width="750px" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow>
                        <asp:TableCell ColumnSpan="3" Width="200px">
                            <asp:Label ID="Label58" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Request
                            Supplies</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="3" Width="550px">
                            <asp:Button ID="btnShowCourierBookingStatus5" runat="server" CausesValidation="False" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell VerticalAlign="Middle" ColumnSpan="6" Width="750px">
                            <asp:Label ID="Label59" runat="server" font-size="XX-Small" font-names="Verdana">Use this page
                            to request your supplies. Once received, your order will actioned immediately and
                            should there be any problem, you will be contacted on the email address below.</asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="170px" ColumnSpan="2">
                            <asp:Label ID="Label60" runat="server" font-size="XX-Small" font-names="Verdana">Contact
                            Name:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="txtContactName" Width="250px"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell Width="250px" ColumnSpan="2"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="170px" ColumnSpan="2">
                            <asp:Label ID="Label61" runat="server" font-size="XX-Small" font-names="Verdana">Contact
                            Email:</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="txtContactEmail" Width="250px"></asp:TextBox>
                        </asp:TableCell>
                        <asp:TableCell Width="250px" ColumnSpan="2"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="170px" ColumnSpan="2">
                            <asp:Label ID="Label62" runat="server" font-size="XX-Small" font-names="Verdana">Contact
                            Telephone:</asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="txtContactTelephone" Width="250px"></asp:TextBox>
                            <br />
                            <br />
                        </asp:TableCell>
                        <asp:TableCell Width="250px" ColumnSpan="2"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px">
                            <asp:Label ID="Label63" runat="server" font-size="XX-Small" font-names="Verdana">Product</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:CheckBox id="chk_LargeFlyerBags" runat="server"></asp:CheckBox>
                            &nbsp;&nbsp;<asp:Label ID="Label64" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Large Flyer
                            Bags(50cm x 40cm)</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="410px" ColumnSpan="2">
                            <asp:Label ID="Label65" runat="server" font-size="XX-Small" font-names="Verdana">Plastic
                            envelopes 50cm x 40cm to protect your documents during transit with a built-in window envelope
                            for your shipping documents.</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:CheckBox id="chk_SmallFlyerBags" runat="server"></asp:CheckBox>
                            &nbsp;&nbsp;<asp:Label ID="Label66" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Small Flyer
                            Bags (40cm x 30cm)</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="110px" HorizontalAlign="Right"></asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                   <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="410px" ColumnSpan="2">
                            <asp:Label ID="Label3a" runat="server" font-size="XX-Small" font-names="Verdana">Plastic
                            envelopes 40cm x 30cm to protect your documents during transit with a built-in window envelope
                            for your shipping documents.</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:CheckBox id="chk_WindowEnvelopes" runat="server"></asp:CheckBox>
                            &nbsp;&nbsp;<asp:Label ID="Label4a" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Window
                            Envelopes</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="110px" HorizontalAlign="Right"></asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="410px" ColumnSpan="2">
                            <asp:Label ID="Label67" runat="server" font-size="XX-Small" font-names="Verdana">Self
                            adhesive envelopes for larger parcels which are used to hold your shipping documents.</asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="140px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="300px">
                            <asp:Label ID="Label68" runat="server" font-size="XX-Small" font-names="Verdana">Click 'Submit'
                            to place your order</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell Width="250px" ColumnSpan="2" HorizontalAlign="Right">
                            <br />
                            <asp:Button ID="btnCancelRequestForSupplies" runat="server" Tooltip="cancel request" onclick="btnCancelRequestForSupplies_Click" Text="cancel" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnSubmitRequestForSupplies" runat="server" onclick="btn_SubmitSuppliesRequest_Click" Tooltip="submit request" Text="submit" />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="60px"></asp:TableCell>
                        <asp:TableCell Width="550px" ColumnSpan="3" HorizontalAlign="Right">
                            <asp:Label id="lblSuppliesMessage" runat="server" forecolor="Red" font-size="XX-Small" font-names="Verdana"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlConfirmDeleteConsignment" runat="server" visible="False" Width="100%">
                <asp:Table id="Table4" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow VerticalAlign="Middle">
                        <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" wrap="False"></asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False">
                            <asp:Button ID="btnShowCourierBookingStatus8" runat="server" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" CausesValidation="false" />
                            &nbsp;
                    </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <asp:Table id="tabConfirmDeleteConsignment" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                    <asp:TableRow>
                        <asp:TableCell wrap="False" HorizontalAlign="Center">
                            <br />
                            <br />
                            <br />
                            <br />
                            <asp:Label ID="Label69" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Are
                            you sure you wish to DELETE Consignment No</asp:Label> &nbsp;<asp:Label runat="server" id="lblConsignmentNumberToDelete" forecolor="Red" font-size="X-Small" font-names="Verdana" font-bold="True"></asp:Label>&nbsp;<asp:Label ID="Label70" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">?</asp:Label>
                            <br />
                            <br />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow >
                        <asp:TableCell HorizontalAlign="Center">
                            <br />
                            <br />
                            <asp:Button ID="btnDeleteConsignment" runat="server" onclick="btn_DeletelConsignment_Click" CausesValidation="False" Tooltip="click yes to delete consignment" Text="yes" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnShowCourierBookingStatus9" runat="server" Tooltip="back to main status page" CausesValidation="False" OnClick="btnShowCourierBookingStatus_Click" Text="go back" />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlConfirmDeleteCollection" runat="server" visible="False" Width="100%">
                <asp:Table id="Table22" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow VerticalAlign="Middle">
                        <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" wrap="False"></asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False">
                            <asp:Button ID="btnShowCourierBookingStatus10" runat="server" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" CausesValidation="false" />
                            &nbsp;
                    </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <asp:Table id="tabConfirmDeleteCollection" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                    <asp:TableRow>
                        <asp:TableCell wrap="False" HorizontalAlign="Center">
                            <br />
                            <br />
                            <br />
                            <br />
                            <asp:Label ID="Label71" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Are
                            you sure you wish to CANCEL Collection No</asp:Label> &nbsp;<asp:Label runat="server" id="lblCollectionNumber" forecolor="Red" font-size="X-Small" font-names="Verdana" font-bold="True"></asp:Label> <asp:Label ID="Label72" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">?</asp:Label>
                            <br />
                            <br />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow >
                        <asp:TableCell HorizontalAlign="Center">
                            <br />
                            <br />
                            <asp:Button ID="btnConfirmCancelCollection" runat="server" onclick="btnConfirmCancelCollection_Click" Tooltip="click yes to confirm you want to cancel this collection" CausesValidation="False" Text="yes" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnShowCourierBookingStatus11" runat="server" CausesValidation="False" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlConfirmNewCourierCollection" runat="server" visible="False" Width="100%">
                <asp:Table id="Table21" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
                    <asp:TableRow VerticalAlign="Middle">
                        <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" wrap="False"></asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False">
                            <asp:Button ID="btnGoBack" runat="server" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="go back" />
                            &nbsp;
                    </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
                <asp:Table id="tabConfirmNewCourierCollection" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                    <asp:TableRow>
                        <asp:TableCell wrap="False" HorizontalAlign="Center">
                            <br />
                            <br />
                            <asp:Label ID="Label74" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Your collection request has been accepted</asp:Label>
                            <br />
                            <br />
                            <asp:Label ID="Label75" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">The
                            Courier Booking Number is</asp:Label>&nbsp;<asp:Label runat="server" id="lblCourierBookingNumber" forecolor="Red" font-size="X-Small" font-names="Verdana" font-bold="True"></asp:Label><asp:Label ID="Label76" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">.</asp:Label>
                            <br />
                            <br />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow >
                        <asp:TableCell HorizontalAlign="Center">
                            <br />
                            <br />
                            <asp:Button ID="btnBookAnotherCollection" runat="server" Tooltip="book another collection" CausesValidation="false" Text="book another collection" OnClick="btnBookACollection_Click" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnShowCourierBookingStatus12" runat="server" CausesValidation="False" Tooltip="back to main status page" OnClick="btnShowCourierBookingStatus_Click" Text="continue" />
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
            <asp:Panel id="pnlConfirmNewConsignment" runat="server" visible="False" Width="100%">
                <asp:Table id="tabConfirmNewConsignment" runat="server" Width="750px" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell ColumnSpan="2" Width="340px">
                            <br/>
                            <asp:Label ID="Label77" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Book a Collection</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="3" Width="380px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="230px"></asp:TableCell>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="170px"></asp:TableCell>
                        <asp:TableCell Width="180px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <br />
                            <br />
                            <asp:Label ID="Label78" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Consignment
                            number</asp:Label> &nbsp;<asp:Label runat="server" id="lblConsignmentNumber" forecolor="Red" font-size="X-Small" font-names="Verdana" font-bold="True"></asp:Label> &nbsp;<asp:Label ID="Label79" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">was
                            successfully added.</asp:Label>
                            <br />
                            <br />
                            <asp:Label id="lblAssociatedWithCollection" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <asp:DataGrid id="grid_1AvailableCollection" runat="server" Width="600px" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True">
                                <FooterStyle wrap="False"></FooterStyle>
                                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                <ItemStyle backcolor="LightGray"></ItemStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="Key" HeaderText="Number">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ReadyAt" HeaderText="Pickup Time" DataFormatString="{0: dd.MM.yy HH:mm}">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CollectionPoint" HeaderText="Collection Point">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ContactName" HeaderText="Contact Name">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Company" HeaderText="Company">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Addr1" HeaderText="Addr1">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Town" HeaderText="City">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                            <asp:DataGrid id="grid_AvailableCollections" runat="server" Width="600px" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="grd_AvailableCollections_item_click">
                                <FooterStyle wrap="False"></FooterStyle>
                                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                <ItemStyle backcolor="LightGray"></ItemStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="Key" HeaderText="Number">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ReadyAt" HeaderText="Pickup Time" DataFormatString="{0: dd.MM.yy HH:mm}">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CollectionPoint" HeaderText="Collection Point">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ContactName" HeaderText="Contact Name">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Company" HeaderText="Company">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Addr1" HeaderText="Addr1">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Town" HeaderText="City">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="Select">
                                        <HeaderStyle font-bold="True" horizontalalign="Center" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle horizontalalign="Center"></ItemStyle>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="CheckBox1" runat="server"></asp:CheckBox>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                            <br />
                            <asp:Label id="lblGridSelection" runat="server" font-size="X-Small" font-names="Verdana" forecolor="Red"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <br />
                            <asp:Label id="lblAssocBookingInstructions" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>
                            <br />
                            <asp:Button ID="btnConsignmentConfirmationYes" runat="server" Tooltip="yes" OnClick="btnConsignmentConfirmationYes_click" Text="yes" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnConsignmentConfirmationNo" runat="server" Tooltip="no" OnClick="btnConsignmentConfirmationNo_click" Text="no" />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>


            <asp:Panel id="pnlAmendCollection" runat="server" visible="False" Width="100%">
                <asp:Table id="tabAmendCollection" runat="server" Width="750px" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Gray">
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell ColumnSpan="2" Width="340px">
                            <br/>
                            <asp:Label ID="Label80" runat="server" font-size="X-Small" font-names="Verdana" font-bold="True">Amend a Collection</asp:Label>
                        </asp:TableCell>
                        <asp:TableCell HorizontalAlign="Right" ColumnSpan="3" Width="380px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="110px"></asp:TableCell>
                        <asp:TableCell Width="230px"></asp:TableCell>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="170px"></asp:TableCell>
                        <asp:TableCell Width="180px"></asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <br />
                            <br />
                            <asp:Label ID="Label81" runat="server" forecolor="Gray" font-size="X-Small" font-names="Verdana" font-bold="True">Consignment
                            number</asp:Label> &nbsp;<asp:Label runat="server" id="lblAmendedConsignment" forecolor="Red" font-size="X-Small" font-names="Verdana" font-bold="True"></asp:Label>
                            <br />
                            <br />
                            <asp:Label id="lblAmendedCollection" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>
                            <br />
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <asp:DataGrid id="grid_OneCollection" runat="server" Width="600px" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True">
                                <FooterStyle wrap="False"></FooterStyle>
                                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                <ItemStyle backcolor="LightGray"></ItemStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="Key" HeaderText="Number">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ReadyAt" HeaderText="Pickup Time" DataFormatString="{0: dd.MM.yy HH:mm}">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CollectionPoint" HeaderText="Collection Point">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ContactName" HeaderText="Contact Name">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Company" HeaderText="Company">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Addr1" HeaderText="Addr1">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Town" HeaderText="City">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                            <asp:DataGrid id="grid_ChooseFromCollections" runat="server" Width="600px" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemCommand="grd_AvailableCollections_item_click">
                                <FooterStyle wrap="False"></FooterStyle>
                                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                <ItemStyle backcolor="LightGray"></ItemStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="Key" HeaderText="Number">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ReadyAt" HeaderText="Pickup Time" DataFormatString="{0: dd.MM.yy HH:mm}">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CollectionPoint" HeaderText="Collection Point">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ContactName" HeaderText="Contact Name">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Company" HeaderText="Company">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Addr1" HeaderText="Addr1">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Town" HeaderText="City">
                                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle wrap="False"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="Select">
                                        <HeaderStyle font-bold="True" horizontalalign="Center" forecolor="Gray"></HeaderStyle>
                                        <ItemStyle horizontalalign="Center"></ItemStyle>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="CheckBox2" runat="server"></asp:CheckBox>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                            <br />
                            <asp:Label id="lblAmendedGridSelection" runat="server" font-size="X-Small" font-names="Verdana" forecolor="Red"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow>
                        <asp:TableCell Width="30px"></asp:TableCell>
                        <asp:TableCell Width="720px" ColumnSpan="5">
                            <asp:Label id="lblAmendBookingInstructions" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana"></asp:Label>
                            <br />
                            <br />
                            <asp:Button ID="btnAmendCollectionCancel" runat="server" Tooltip="cancel" OnClick="btnAmendCollectionCancel_Click" Text="cancel" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnAmendCollectionSubmit" runat="server" onclick="btnAmendCollectionSubmit_Click" Tooltip="submit" Text="submit" />
                            &nbsp;&nbsp;
                            <asp:Button ID="btnUnBookConsignment" runat="server" onclick="btnUnBookConsignment_Click" Tooltip="unbook consignment" Text="unbook" />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>

            <asp:Label id="lblError" runat="server" forecolor="#00C000" font-names="Verdana" font-size="X-Small"></asp:Label>
        <script type="text/javascript">
           function PrintConsignment(Key){
                window.open("ConsignmentNote.aspx?key=" + Key,"Consignment","top=10,left=10,width=650,height=480,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=no");
           }
           function PrintLabel(Key){
                window.open("consignment_label.aspx?key=" + Key,"Consignment","top=10,left=10,width=350,height=420,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=no");
           }
           function PrintManifest(Key){
                window.open("DriversManifest.aspx?key=" + Key,"Manifest","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");
           }
           function chkDocs(evt){
                var e = evt || event;
                var elm = e.target || e.srcElement;
                if (elm.value != 'Documents') {alert("You have changed the Description field from the default 'Documents'.\n\nIf your consignment includes one or more non-document items please ensure the Docs only check box has been changed to indicate this.\n\nThank you.");}
           }
        </script>
    </form>
</body>
</html>
