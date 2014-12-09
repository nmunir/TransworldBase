<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data " %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Globalization" %>
<%@ import Namespace="System.Threading" %>
<%@ import Namespace="System.Collections.Generic" %>
        
<script runat="server">

    ' SPROCS
    ' spASPNET_Address_GetUserAddresses
    ' spASPNET_Address_GetGlobalAddresses
    ' spASPNET_Country_GetCountries
    ' spASPNET_Address_GetAddressFromKey
    ' spASPNET_Address_Add
    ' spASPNET_Address_AddToGlobal
    ' spASPNET_Address_AddToPersonal
    ' spASPNET_Address_Update
    ' spASPNET_Address_DeletePersonal
    ' spASPNET_Address_DeleteGlobal
    
    ' spASPNET_Address_GetAddressesPaged

    ' TO DO
    ' add dropdown to select page size
    ' use DataReader instead of DataAdapter
    ' handle diacritic characters when exporting address book
    ' check page boundaries
    ' COUNTRY cannot be edited
    ' after editing an address grid is not refreshed automatically
    ' check tabbing order & keyboard navigability
    ' incorporate upload utility
    ' convert ASP tables to standard tables
    ' fix export
    ' fix transaction
    ' put common address routines into external class to share with Order page
    ' need to check why, when a new address is created, an entry is inserted into the UserAddressBookProfile table by the stored procedure
    ' Upload did not recognise CHINA, PEOPLES REP when I put it in explicitly

    ' TO DO ON DISTRIBUTION LISTS
    ' Norwegian addresses (see Firda Sj?farmer As) not handled well
    ' when search string not found, no error message returned
    ' renaming with apostrophe as character 50 causes invalid SQL expression
    ' check whether delete flag should be observed in distribution lists
    ' check handling of EditGAB when ViewGAB is false
    
    Const CUSTOMER_WU As Integer = 579
    Const CUSTOMER_WUIRE As Integer = 686
    
    Const COUNTRY_CODE_CANADA As Int32 = 38
    Const COUNTRY_CODE_USA As Int32 = 223
    Const COUNTRY_CODE_USA_NYC As Int32 = 256


    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Const START_PAGE As Integer = 0
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call GetSiteFeatures()
            If Session("ViewGAB") = False Then
                btnViewSharedAddressBook.Visible = False
                btnDistributionLists1.Visible = False
                btnDistributionLists2.Visible = False
            Else
                btnViewSharedAddressBook.Visible = True
                btnDistributionLists1.Visible = True
                btnDistributionLists2.Visible = True
            End If
            If Session("EditGAB") = False Then
                dgSharedAddressBook.Columns(7).Visible = False
                dgPersonalAddressBook.Columns(8).Visible = False
                btnSharedAddressBookAddNewAddress.Visible = False
                btnNewDistributionList.Visible = False
                btnFinishEditingList.Text = "back to distribution lists"
                lblInstructions.Visible = False
                tblEditDistributionList.Visible = False
                btnRenameThisDistributionList.Visible = False
                btnDeleteThisDistributionList.Visible = False
            End If
            
            Call GetCountries()
            Call ShowPersonalAddressBookPanel()
            pbIsEditingDistributionList = False

        End If
        txtPersonalAddressBookSearchCriteria.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchAddress.ClientID + "')")
        txtSharedAddressBookSearchCriteria.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchSharedAddressBook.ClientID + "')")
        tbSharedAddressBookSearchCriteriaForDistbnLists.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchSharedAddressBookForDistbnLists.ClientID + "')")
        tbNewDistributionListName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnRename.ClientID + "')")

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)
        Call SetTitle()
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Address Book"
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
        pbProductOwners = dr("ProductOwners")
        btnDistributionLists1.Visible = pbProductOwners
        btnDistributionLists2.Visible = pbProductOwners
    End Sub
    
    Protected Function IsWUorWUIRE() As Boolean
        Dim arrCustomerFEXCO() As Integer = {CUSTOMER_WU, CUSTOMER_WUIRE}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWUorWUIRE = IIf(gbSiteTypeDefined, gsSiteType = "fexco", Array.IndexOf(arrCustomerFEXCO, nCustomerKey) >= 0)
    End Function

    Protected Sub HideAllPanels()
        pnlAddressDetail.Visible = False
        pnlSharedAddressBookList.Visible = False
        pnlPersonalAddressBook.Visible = False
        pnlDistributionLists.Visible = False
        pnlDistributionList.Visible = False
        pnlRenameDistributionList.Visible = False
    End Sub
    
    Protected Sub ShowPersonalAddressBookPanel()
        Call HideAllPanels()
        pnlPersonalAddressBook.Visible = True
        lblAddressBook.Text = "<b>Viewing:</b> personal address book"
        pbIsViewingGlobal = False
        pbIsEditingDistributionList = False
    End Sub
    
    Protected Sub ShowAddressDetailPanel()
        Call HideAllPanels()
        pnlAddressDetail.Visible = True
    End Sub
    
    Protected Sub ShowDistributionListsPanel()
        Call HideAllPanels()
        pnlDistributionLists.Visible = True
        lblAddressBook.Text = "Distribution Lists"
        pbIsEditingDistributionList = True
    End Sub
    
    Protected Sub ShowDistributionListPanel()
        Call HideAllPanels()
        pnlDistributionList.Visible = True
        lblAddressBook.Text = "Editing Distribution List  <b>" & psDistributionList & "</b>"
        pbIsEditingDistributionList = True
    End Sub
    
    Protected Sub ShowGlobalAddressPanel()
        Call HideAllPanels()
        pnlSharedAddressBookList.Visible = True
        lblAddressBook.Text = "<b>Viewing:</b> shared address book"
        pbIsViewingGlobal = True
        pbIsEditingDistributionList = False
    End Sub
    
    Protected Sub dgPersonalAddressBook_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "info" Then
            Dim itemCell As TableCell = e.Item.Cells(0)
            plAddressKey = CLng(itemCell.Text)
            pbIsAddingAddress = False
            GetConsigneeAddress(plAddressKey)
            ShowAddressDetailPanel()
        End If
    End Sub
    
    Protected Sub dgSharedAddressBook_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "info" Then
            Dim itemCell As TableCell = e.Item.Cells(0)
            plAddressKey = CLng(itemCell.Text)
            pbIsAddingAddress = False
            GetConsigneeAddress(plAddressKey)
            ShowAddressDetailPanel()
        End If
    End Sub
    
    Protected Sub InitDataGrids()
        pnPage = START_PAGE
        dgPersonalAddressBook.CurrentPageIndex = 0
        dgSharedAddressBook.CurrentPageIndex = 0
        dgSharedAddressBookForDistbnLists.CurrentPageIndex = 0
        pnVirtualItemCount = nGetAddressCount()
    End Sub
    
    Protected Function nGetAddressCount() As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim dtRecordCount As New DataTable
        Dim nRecordCount As Integer
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddressCount", oConn)
        Dim sSearchCriteria As String
        If pbIsEditingDistributionList Then
            sSearchCriteria = tbSharedAddressBookSearchCriteriaForDistbnLists.Text
        Else
            If pbIsViewingGlobal Then
                sSearchCriteria = txtSharedAddressBookSearchCriteria.Text
            Else
                sSearchCriteria = txtPersonalAddressBookSearchCriteria.Text
            End If
        End If
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            If pbIsViewingGlobal Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
            Else
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            End If
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            
            oAdapter.Fill(dtRecordCount)
            nRecordCount = CInt(dtRecordCount.Rows(0).Item(0))
            nGetAddressCount = nRecordCount
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Function ReadPage() As DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddressesPaged", oConn)
        Dim sSearchCriteria As String
        If pbIsEditingDistributionList Then
            sSearchCriteria = tbSharedAddressBookSearchCriteriaForDistbnLists.Text
        Else
            If pbIsViewingGlobal Then
                sSearchCriteria = txtSharedAddressBookSearchCriteria.Text
            Else
                sSearchCriteria = txtPersonalAddressBookSearchCriteria.Text
            End If
        End If
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            
            Dim nPageStart As Integer = ((pnPage) * 20) + 1
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@PageStart", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@PageStart").Value = nPageStart
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@RowsToReturn", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@RowsToReturn").Value = 20
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            If pbIsViewingGlobal Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
            Else
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            End If
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SortKey", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SortKey").Value = psSortExpression

            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        ReadPage = oDataTable
    End Function
    
    Protected Sub BindPersonalAddressGrid_PAGINGVERSON()
        Dim oDataTable As DataTable = ReadPage()
        lblAddressMessage.Text = ""
        If oDataTable.Rows.Count > 0 Then
            btnRefreshAddressGrid.Visible = True
            btnExportPersonalAddressBook.Visible = True
            dgPersonalAddressBook.Visible = True
            dgPersonalAddressBook.DataSource = oDataTable
            dgPersonalAddressBook.VirtualItemCount = pnVirtualItemCount
            dgPersonalAddressBook.DataBind()
            btnRefreshAddressGrid.Visible = True
        Else
            btnRefreshAddressGrid.Visible = False
            btnExportPersonalAddressBook.Visible = False
            dgPersonalAddressBook.Visible = False
            lblAddressMessage.Text = "No matching records"
            btnRefreshAddressGrid.Visible = False
        End If
        ShowPersonalAddressBookPanel()
    End Sub

    Protected Sub BindPersonalAddressGrid() ' NON PAGING VERSION
        Dim sSearchCriteria As String
        If pbIsEditingDistributionList Then
            sSearchCriteria = tbSharedAddressBookSearchCriteriaForDistbnLists.Text
        Else
            If pbIsViewingGlobal Then
                sSearchCriteria = txtSharedAddressBookSearchCriteria.Text
            Else
                sSearchCriteria = txtPersonalAddressBookSearchCriteria.Text
            End If
        End If

        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetUserAddresses", oConn)
        ' Dim sSearchCriteria As String = txtAddressSearchCriteria.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            If Session("CustomerKey") > 0 Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            Else
                'If we're an Acct Handler we need to get the GAB, so tell the proc we're not
                'a customer user
                oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
            End If
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = 0
            oAdapter.Fill(oDataSet, "Addresses")
            Dim Source As DataView = oDataSet.Tables("Addresses").DefaultView
            Source.Sort = psSortExpression
            If Source.Count > 0 Then
                btnRefreshAddressGrid.Visible = True
                btnExportPersonalAddressBook.Visible = True
                dgPersonalAddressBook.Visible = True
                dgPersonalAddressBook.DataSource = Source
                dgPersonalAddressBook.DataBind()
                btnRefreshAddressGrid.Visible = True
            Else
                btnRefreshAddressGrid.Visible = False
                btnExportPersonalAddressBook.Visible = False
                dgPersonalAddressBook.Visible = False
                lblAddressMessage.Text = "No matching records"
                btnRefreshAddressGrid.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call ShowPersonalAddressBookPanel()
    End Sub

    Protected Sub BindSharedAddressGrid_PAGINGVERSON()
        Dim oDataTable As DataTable = ReadPage()
        lblSharedAddressBookMessage.Text = ""
        If oDataTable.Rows.Count > 0 Then
            btnRefreshSharedAddressBook.Visible = True
            btnExportSharedAddressBook.Visible = True
            dgSharedAddressBook.Visible = True
            dgSharedAddressBook.DataSource = oDataTable
            dgSharedAddressBook.VirtualItemCount = pnVirtualItemCount
            dgSharedAddressBook.DataBind()
        Else
            btnRefreshSharedAddressBook.Visible = False
            btnExportSharedAddressBook.Visible = False
            dgSharedAddressBook.Visible = False
            lblSharedAddressBookMessage.Text = "No matching records"
        End If
        ShowGlobalAddressPanel()
    End Sub

    Protected Sub BindSharedAddressGrid() ' NON PAGING VERSION
        Dim sSearchCriteria As String
        If pbIsEditingDistributionList Then
            sSearchCriteria = tbSharedAddressBookSearchCriteriaForDistbnLists.Text
        Else
            If pbIsViewingGlobal Then
                sSearchCriteria = txtSharedAddressBookSearchCriteria.Text
            Else
                sSearchCriteria = txtPersonalAddressBookSearchCriteria.Text
            End If
        End If

        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetGlobalAddresses", oConn)
        ' Dim sSearchCriteria As String = txtGABSearchCriteria.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblSharedAddressBookMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.Fill(oDataSet, "Addresses")
            Dim Source As DataView = oDataSet.Tables("Addresses").DefaultView
            Source.Sort = psSortExpression
            If Source.Count > 0 Then
                btnRefreshSharedAddressBook.Visible = True
                btnExportSharedAddressBook.Visible = True
                dgSharedAddressBook.Visible = True
                dgSharedAddressBook.DataSource = Source
                dgSharedAddressBook.DataBind()
            Else
                btnRefreshSharedAddressBook.Visible = False
                btnExportSharedAddressBook.Visible = False
                dgSharedAddressBook.Visible = False
                lblSharedAddressBookMessage.Text = "No matching records"
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        ShowGlobalAddressPanel()
    End Sub

    Protected Sub BindSharedAddressGridForDistbnLists()
        Dim oDataTable As DataTable = ReadPage()
        If oDataTable.Rows.Count > 0 Then
            dgSharedAddressBookForDistbnLists.Visible = True
            dgSharedAddressBookForDistbnLists.DataSource = oDataTable
            dgSharedAddressBookForDistbnLists.VirtualItemCount = pnVirtualItemCount
            dgSharedAddressBookForDistbnLists.DataBind()
        Else
            dgSharedAddressBookForDistbnLists.Visible = False
            lblSharedAddressBookMessage.Text = "No matching records"
        End If
    End Sub
    
    Protected Sub ExportPersonalAddressBook()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetUserAddresses", oConn)
        Dim sSearchCriteria As String = txtPersonalAddressBookSearchCriteria.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.Fill(oDataSet, "Addresses")
            Dim Source As DataView = oDataSet.Tables("Addresses").DefaultView
            'Source.Sort = SortField
    
            If Source.Count > 0 Then
                Response.Clear()
                'Response.ContentType = "Application/x-msexcel"
                Response.ContentType = "text/csv"
                Response.AddHeader("Content-Disposition", "attachment; filename=address_book.csv")
    
                Dim r As DataRowView
                Dim c As DataColumn
                Dim sItem As String
    
                Dim IgnoredItems As New ArrayList
    
                IgnoredItems.Add("DestKey")
                IgnoredItems.Add("CountryKey")
                IgnoredItems.Add("DefaultCommodityId")
                IgnoredItems.Add("DefaultSpecialInstructions")
                IgnoredItems.Add("Fax")
                IgnoredItems.Add("Email")
    
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
            End If
    
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    
    End Sub
    
    Protected Sub ExportSharedAddressBook()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetGlobalAddresses", oConn)
        Dim sSearchCriteria As String = txtSharedAddressBookSearchCriteria.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblSharedAddressBookMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.Fill(oDataSet, "Addresses")
            Dim Source As DataView = oDataSet.Tables("Addresses").DefaultView
            'Source.Sort = SortField
    
            If Source.Count > 0 Then
                Response.Clear()
                'Response.ContentType = "Application/x-msexcel"
                Response.ContentType = "text/csv"
                Response.AddHeader("Content-Disposition", "attachment; filename=address_book.csv")
                Dim r As DataRowView
                Dim c As DataColumn
                Dim sItem As String
                Dim IgnoredItems As New ArrayList
                IgnoredItems.Add("DestKey")
                IgnoredItems.Add("CountryKey")
                IgnoredItems.Add("DefaultCommodityId")
                IgnoredItems.Add("DefaultSpecialInstructions")
                IgnoredItems.Add("Fax")
                IgnoredItems.Add("Email")
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
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub dgPersonalAddressBook_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs)
        Dim dgscea As DataGridSortCommandEventArgs = e
        Call InitDataGrids()
        psSortExpression = e.SortExpression
        Call BindPersonalAddressGrid()
    End Sub
    
    Protected Sub dgSharedAddressBook_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs)
        Dim dgscea As DataGridSortCommandEventArgs = e
        Call InitDataGrids()
        psSortExpression = e.SortExpression
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub dgSharedAddressBookForDistbnLists_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs)
        Dim dgscea As DataGridSortCommandEventArgs = e
        Call InitDataGrids()
        psSortExpression = e.SortExpression
        Call BindSharedAddressGridForDistbnLists()
    End Sub
    
    Protected Sub GetCountries()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Country_GetCountries", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            ddlCountry.DataSource = oCmd.ExecuteReader()
            ddlCountry.DataTextField = "CountryName"
            ddlCountry.DataValueField = "CountryKey"
            ddlCountry.DataBind()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub GetConsigneeAddress(ByVal plAddressKey As Long)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Address_GetAddressFromKey2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        'Dim oParam As New SqlParameter("@DestKey", SqlDbType.Int, 4)
        Dim oParam As New SqlParameter("@AddressKey", SqlDbType.Int, 4)
        oCmd.Parameters.Add(oParam)
        oParam.Value = plAddressKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If IsDBNull(oDataReader("Code")) Then
                txtCode.Text = ""
            Else
                txtCode.Text = oDataReader("Code")
            End If
            If IsDBNull(oDataReader("Company")) Then
                txtCompany.Text = ""
            Else
                txtCompany.Text = oDataReader("Company")
            End If
            If IsDBNull(oDataReader("Addr1")) Then
                txtAddr1.Text = ""
            Else
                txtAddr1.Text = oDataReader("Addr1")
            End If
            If IsDBNull(oDataReader("Addr2")) Then
                txtAddr2.Text = ""
            Else
                txtAddr2.Text = oDataReader("Addr2")
            End If
            If IsDBNull(oDataReader("Addr3")) Then
                txtAddr3.Text = ""
            Else
                txtAddr3.Text = oDataReader("Addr3")
            End If
            If IsDBNull(oDataReader("Town")) Then
                txtCity.Text = ""
            Else
                txtCity.Text = oDataReader("Town")
            End If
            If IsDBNull(oDataReader("CountryName")) Then
                ddlCountry.SelectedItem.Text = ""
                ddlCountry.SelectedItem.Value = 0
            Else
                ddlCountry.SelectedItem.Value = oDataReader("CountryKey")
                plCountryKey = oDataReader("CountryKey")
                ddlCountry.SelectedItem.Text = oDataReader("CountryName")
                txtCountry.Text = oDataReader("CountryName")
            End If
            If IsDBNull(oDataReader("State")) Then
                txtState.Text = ""
            Else
                txtState.Text = oDataReader("State")
            End If

            Call SetCountry(plCountryKey, txtState.Text)
            
            If IsDBNull(oDataReader("PostCode")) Then
                txtPostCode.Text = ""
            Else
                txtPostCode.Text = oDataReader("PostCode")
            End If
            If IsDBNull(oDataReader("AttnOf")) Then
                txtAttnOf.Text = ""
            Else
                txtAttnOf.Text = oDataReader("AttnOf")
            End If
            If IsDBNull(oDataReader("Telephone")) Then
                txtTel.Text = ""
            Else
                txtTel.Text = oDataReader("Telephone")
            End If
            oDataReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub AddNewAddress()
        Dim bError As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_Add", oConn)
        Dim oTrans As SqlTransaction
        oCmd.CommandType = CommandType.StoredProcedure
        plAddressKey = -1        'the address key is returned by proc making insert so initialise it here
        plCountryKey = CLng(ddlCountry.SelectedItem.Value)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramCode As SqlParameter = New SqlParameter("@Code", SqlDbType.NVarChar, 20)
        paramCode.Value = txtCode.Text
        oCmd.Parameters.Add(paramCode)
        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = txtCompany.Text
        oCmd.Parameters.Add(paramCompany)
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = txtAddr1.Text
        oCmd.Parameters.Add(paramAddr1)
        Dim paramparamAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramparamAddr2.Value = txtAddr2.Text
        oCmd.Parameters.Add(paramparamAddr2)
        Dim paramparamAddr3 As SqlParameter = New SqlParameter("@Addr3", SqlDbType.NVarChar, 50)
        paramparamAddr3.Value = txtAddr3.Text
        oCmd.Parameters.Add(paramparamAddr3)
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = txtCity.Text
        oCmd.Parameters.Add(paramTown)
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        ' paramState.Value = txtState.Text
        paramState.Value = GetState()
        oCmd.Parameters.Add(paramState)
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = txtPostCode.Text
        oCmd.Parameters.Add(paramPostCode)
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = plCountryKey
        oCmd.Parameters.Add(paramCountryKey)
        Dim paramDefaultCommodityId As SqlParameter = New SqlParameter("@DefaultCommodityId", SqlDbType.NVarChar, 100)
        paramDefaultCommodityId.Value = ""
        oCmd.Parameters.Add(paramDefaultCommodityId)
        Dim paramDefaultSpecialInstructions As SqlParameter = New SqlParameter("@DefaultSpecialInstructions", SqlDbType.NVarChar, 100)
        paramDefaultSpecialInstructions.Value = ""
        oCmd.Parameters.Add(paramDefaultSpecialInstructions)
        Dim paramAttnOf As SqlParameter = New SqlParameter("@AttnOf", SqlDbType.NVarChar, 50)
        paramAttnOf.Value = txtAttnOf.Text
        oCmd.Parameters.Add(paramAttnOf)
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 50)
        paramTelephone.Value = txtTel.Text
        oCmd.Parameters.Add(paramTelephone)
        Dim paramFax As SqlParameter = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        paramFax.Value = ""
        oCmd.Parameters.Add(paramFax)
        Dim paramEmail As SqlParameter = New SqlParameter("@Email", SqlDbType.NVarChar, 50)
        paramEmail.Value = ""
        oCmd.Parameters.Add(paramEmail)
        Dim paramLastUpdatedByKey As SqlParameter = New SqlParameter("@LastUpdatedByKey", SqlDbType.Int, 4)
        paramLastUpdatedByKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramLastUpdatedByKey)
        Dim paramAddressKey As SqlParameter = New SqlParameter("@AddressKey", SqlDbType.Int, 4)
        paramAddressKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramAddressKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
            plAddressKey = paramAddressKey.Value
        Catch ex As SqlException
            oTrans.Rollback("AddRecord")
            bError = True
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub AddToGlobalAddressBook()
        If plAddressKey > 0 Then
            Dim bError As Boolean = False
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToGlobal", oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            ' Add Parameters to SPROC
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            paramCustomerKey.Value = Session("CustomerKey")
            oCmd.Parameters.Add(paramCustomerKey)
            ' Add Parameters to SPROC
            Dim paramAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramAddressKey.Value = plAddressKey
            oCmd.Parameters.Add(paramAddressKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                oCmd.Connection = oConn
                oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                oTrans.Commit()
            Catch ex As SqlException
                oTrans.Rollback("AddRecord")
                bError = True
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            If Not bError Then
                lblAddToSharedAddressBook.Text = "Address added to Shared Address Book"
            End If
        End If
    End Sub
    
    Protected Sub AddToPersonalAddressBook()
        If plAddressKey > 0 Then
            Dim bError As Boolean = False
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToPersonal", oConn)
            Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramCustomerKey.Value = Session("UserKey")
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramAddressKey.Value = plAddressKey
            oCmd.Parameters.Add(paramAddressKey)
            Try
                oConn.Open()
                oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                oCmd.Connection = oConn
                oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                oTrans.Commit()
            Catch ex As SqlException
                oTrans.Rollback("AddRecord")
                bError = True
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            If Not bError Then
                'btnAddToSharedAddressBook.Visible = True              'prevent same address being added twice after successful add
                lblAddToPersonalAddressBook.Text = "Address added to Personal Address Book"
            End If
        Else
            ' cannot insert without a valid address key
        End If
    End Sub
    
    Protected Sub CopyToPersonalAddressBook()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi In dgSharedAddressBook.Items
            cb = CType(dgi.Cells(8).Controls(1), CheckBox)
            If cb.Checked Then
                Dim celplAddressKey As TableCell = dgi.Cells(0)
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToPersonal", oConn)
                Dim oTrans As SqlTransaction
                oCmd.CommandType = CommandType.StoredProcedure
                ' Add Parameters to SPROC
                Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                paramUserKey.Value = Session("UserKey")
                oCmd.Parameters.Add(paramUserKey)
                Dim paramGlobalAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
                paramGlobalAddressKey.Value = CLng(celplAddressKey.Text)
                oCmd.Parameters.Add(paramGlobalAddressKey)
                Try
                    oConn.Open()
                    oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                    oCmd.Connection = oConn
                    oCmd.Transaction = oTrans
                    oCmd.ExecuteNonQuery()
                    oTrans.Commit()
                Catch ex As SqlException
                    oTrans.Rollback("AddRecord")
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
            End If
        Next dgi
    End Sub
    
    Protected Sub CopyToSharedAddressBook()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi In dgPersonalAddressBook.Items
            cb = CType(dgi.Cells(8).Controls(1), CheckBox)
            If cb.Checked Then
                Dim celplAddressKey As TableCell = dgi.Cells(0)
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToGlobal", oConn)
                Dim oTrans As SqlTransaction
                oCmd.CommandType = CommandType.StoredProcedure
                ' Add Parameters to SPROC
                Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                paramCustomerKey.Value = Session("CustomerKey")
                oCmd.Parameters.Add(paramCustomerKey)
                Dim paramGlobalAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
                paramGlobalAddressKey.Value = CLng(celplAddressKey.Text)
                oCmd.Parameters.Add(paramGlobalAddressKey)
                Try
                    oConn.Open()
                    oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddRecord")
                    oCmd.Connection = oConn
                    oCmd.Transaction = oTrans
                    oCmd.ExecuteNonQuery()
                    oTrans.Commit()
                Catch ex As SqlException
                    oTrans.Rollback("AddRecord")
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
            End If
        Next dgi
    End Sub
    
    Protected Sub UpdateExistingAddress()
        Dim bError As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_Update", oConn)
        Dim oTrans As SqlTransaction
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramKey As SqlParameter = New SqlParameter("@Key", SqlDbType.Int, 4)
        paramKey.Value = plAddressKey
        oCmd.Parameters.Add(paramKey)
        Dim paramCode As SqlParameter = New SqlParameter("@Code", SqlDbType.NVarChar, 20)
        paramCode.Value = txtCode.Text
        oCmd.Parameters.Add(paramCode)
        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = txtCompany.Text
        oCmd.Parameters.Add(paramCompany)
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = txtAddr1.Text
        oCmd.Parameters.Add(paramAddr1)
        Dim paramparamAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramparamAddr2.Value = txtAddr2.Text
        oCmd.Parameters.Add(paramparamAddr2)
        Dim paramparamAddr3 As SqlParameter = New SqlParameter("@Addr3", SqlDbType.NVarChar, 50)
        paramparamAddr3.Value = txtAddr3.Text
        oCmd.Parameters.Add(paramparamAddr3)
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = txtCity.Text
        oCmd.Parameters.Add(paramTown)
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        ' paramState.Value = txtState.Text
        paramState.Value = GetState()
        oCmd.Parameters.Add(paramState)
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = txtPostCode.Text
        oCmd.Parameters.Add(paramPostCode)
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = plCountryKey
        oCmd.Parameters.Add(paramCountryKey)
        Dim paramDefaultCommodityId As SqlParameter = New SqlParameter("@DefaultCommodityId", SqlDbType.NVarChar, 100)
        paramDefaultCommodityId.Value = ""
        oCmd.Parameters.Add(paramDefaultCommodityId)
        Dim paramDefaultSpecialInstructions As SqlParameter = New SqlParameter("@DefaultSpecialInstructions", SqlDbType.NVarChar, 100)
        paramDefaultSpecialInstructions.Value = ""
        oCmd.Parameters.Add(paramDefaultSpecialInstructions)
        Dim paramAttnOf As SqlParameter = New SqlParameter("@AttnOf", SqlDbType.NVarChar, 50)
        paramAttnOf.Value = txtAttnOf.Text
        oCmd.Parameters.Add(paramAttnOf)
        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 50)
        paramTelephone.Value = txtTel.Text
        oCmd.Parameters.Add(paramTelephone)
        Dim paramFax As SqlParameter = New SqlParameter("@Fax", SqlDbType.NVarChar, 50)
        paramFax.Value = ""
        oCmd.Parameters.Add(paramFax)
        Dim paramEmail As SqlParameter = New SqlParameter("@Email", SqlDbType.NVarChar, 50)
        paramEmail.Value = ""
        oCmd.Parameters.Add(paramEmail)
        Dim paramLastUpdatedByKey As SqlParameter = New SqlParameter("@LastUpdatedByKey", SqlDbType.Int, 4)
        paramLastUpdatedByKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramLastUpdatedByKey)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "UpdateRecord")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
        Catch ex As SqlException
            oTrans.Rollback("UpdateRecord")
            bError = True
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
            If bError = False Then
                Call ShowPersonalAddressBookPanel()
            End If
        End Try
    End Sub
    
    Protected Function GetState() As String
        If txtState.Visible Then
            GetState = txtState.Text
        ElseIf ddlUSStatesCanadianProvinces.Visible Then
            GetState = ddlUSStatesCanadianProvinces.SelectedItem.Text
        ElseIf lblLegendNewYorkCity.Visible Then
            GetState = lblLegendNewYorkCity.Text
        Else
            Call WebMsgBox.Show("Could not identify source of state - please contact development.")
            GetState = "XX"
        End If
    End Function
    
    Protected Sub DeleteFromPersonalAddressBook()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi In dgPersonalAddressBook.Items
            cb = CType(dgi.Cells(7).Controls(1), CheckBox)
            If cb.Checked Then
                Dim bError As Boolean = False
                Dim celplAddressKey As TableCell = dgi.Cells(0)
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_DeletePersonal", oConn)
                Dim oTrans As SqlTransaction
                oCmd.CommandType = CommandType.StoredProcedure
                Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                paramUserKey.Value = Session("UserKey")
                oCmd.Parameters.Add(paramUserKey)
                Dim paramGlobalAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
                paramGlobalAddressKey.Value = CLng(celplAddressKey.Text)
                oCmd.Parameters.Add(paramGlobalAddressKey)
                Try
                    oConn.Open()
                    oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "DeleteRecord")
                    oCmd.Connection = oConn
                    oCmd.Transaction = oTrans
                    oCmd.ExecuteNonQuery()
                    oTrans.Commit()
                Catch ex As SqlException
                    oTrans.Rollback("DeleteRecord")
                    bError = True
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                    If bError = False Then
                        ShowPersonalAddressBookPanel()
                    End If
                End Try
            End If
        Next dgi
    End Sub
    
    Protected Sub DeleteFromSharedAddressBook()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi In dgSharedAddressBook.Items
            cb = CType(dgi.Cells(7).Controls(1), CheckBox)
            If cb.Checked Then
                Dim celplAddressKey As TableCell = dgi.Cells(0)
                Dim oConn As New SqlConnection(gsConn)
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_DeleteGlobal", oConn)
                Dim oTrans As SqlTransaction
                oCmd.CommandType = CommandType.StoredProcedure
                ' Add Parameters to SPROC
                Dim paramUserKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                paramUserKey.Value = Session("CustomerKey")
                oCmd.Parameters.Add(paramUserKey)
                Dim paramGlobalAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
                paramGlobalAddressKey.Value = CLng(celplAddressKey.Text)
                oCmd.Parameters.Add(paramGlobalAddressKey)
                Try
                    oConn.Open()
                    oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "DeleteRecord")
                    oCmd.Connection = oConn
                    oCmd.Transaction = oTrans
                    oCmd.ExecuteNonQuery()
                    oTrans.Commit()
                Catch ex As SqlException
                    oTrans.Rollback("DeleteRecord")
                    lblError.Text = ex.ToString
                Finally
                    oConn.Close()
                End Try
            End If
        Next dgi
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub ResetForm()
        txtCode.Text = ""
        txtCompany.Text = ""
        txtAddr1.Text = ""
        txtAddr2.Text = ""
        txtAddr3.Text = ""
        txtCity.Text = ""
        txtState.Text = ""
        txtPostCode.Text = ""
        ddlCountry.SelectedIndex = -1
        txtAttnOf.Text = ""
        txtTel.Text = ""
        plCountryKey = "-1"
        plAddressKey = "-1"
        lblAddToPersonalAddressBook.Text = ""
        lblAddToSharedAddressBook.Text = ""
        Call SetCountryOther()
    End Sub
    
    Protected Sub btnRefreshAddressGrid_click(ByVal s As Object, ByVal e As EventArgs)
        psSortExpression = "CountryName"
        Call BindPersonalAddressGrid()
    End Sub
    
    Protected Sub btnRefreshSharedAddressBook_click(ByVal s As Object, ByVal e As EventArgs)
        psSortExpression = "CountryName"
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub btn_SearchForAddress_click(ByVal sender As Object, ByVal e As ImageClickEventArgs)
        psSortExpression = "CountryName"
        Call InitDataGrids()
        Call BindPersonalAddressGrid()
    End Sub
    
    Protected Sub btnShowFullAddressList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtPersonalAddressBookSearchCriteria.Text = ""
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindPersonalAddressGrid()
    End Sub

    Protected Sub btnSearchAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindPersonalAddressGrid()
    End Sub
    
    Protected Sub InitForNewAddress()
        pbIsAddingAddress = True
        If IsWUorWUIRE() Then
            txtAttnOf.Text = "Western Union Operator"
        End If
        txtCompany.Focus()
    End Sub
    
    Protected Sub btnPersonalAddressbookAddNewAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitForNewAddress()
        Call ShowAddressDetailPanel()
    End Sub
    
    Protected Sub btnExportPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs )
        Call ExportPersonalAddressBook()
    End Sub
    
    Protected Sub btnViewSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowGlobalAddressPanel()
    End Sub
    
    Protected Sub btnDeleteFromPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DeleteFromPersonalAddressBook()
        Call BindPersonalAddressGrid()
    End Sub
    
    Protected Sub btnCopyToSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CopyToSharedAddressBook()
    End Sub
    
    Protected Sub btnAddToPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddNewAddress()
        Call AddToPersonalAddressBook()
        btnAddToPersonalAddressBook.Visible = False                ' prevent same address being added twice after successful add
    End Sub
    
    Protected Sub btnAddToSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddNewAddress()
        Call AddToGlobalAddressBook()            '
        btnAddToSharedAddressBook.Visible = False     'prevent same address being added twice after successful add
    End Sub

    Protected Sub btnAddAnotherAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ResetForm()
        If pbIsViewingGlobal Then
            btnAddToSharedAddressBook.Visible = True
        Else
            btnAddToPersonalAddressBook.Visible = True
            If CBool(Session("EditGAB")) Then
                btnAddToSharedAddressBook.Visible = True
            Else
                btnAddToSharedAddressBook.Visible = False
            End If
        End If
    End Sub
    
    Protected Sub btnCloseAddressDetailPanelFromAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CloseAddressDetailPanel()
    End Sub
    
    Protected Sub CloseAddressDetailPanel()
        Call ResetForm()
        plAddressKey = -1
        If Not pbIsViewingGlobal Then
            btnAddToSharedAddressBook.Visible = False
        End If
        btnAddToPersonalAddressBook.Visible = True
        If pbIsViewingGlobal Then
            Call ShowGlobalAddressPanel()
        Else
            Call ShowPersonalAddressBookPanel()
        End If
    End Sub
    
    Protected Sub btnCloseAddressDetailPanelFromEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CloseAddressDetailPanel()
    End Sub
    
    Protected Sub btnCloseAddressDetailPanelFromView_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CloseAddressDetailPanel()
    End Sub
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call UpdateExistingAddress()
    End Sub
    
    Protected Sub btnShowSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSharedAddressBookSearchCriteria.Text = ""
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub btnSearchSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub btnSearchSharedAddressBookForDistbnLists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pbIsViewingGlobal = True
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindSharedAddressGridForDistbnLists()
        lblInstructions.Text = String.Empty
        lblInstructions.Visible = False
    End Sub
    
    Protected Sub btnSharedAddressBookAddNewAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitForNewAddress()
        Call ShowAddressDetailPanel()
    End Sub
    
    Protected Sub btnExportSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs )
        psSortExpression = "CountryName"
        Call ExportSharedAddressBook()
    End Sub
    
    Protected Sub btnViewPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs )
        Call ShowPersonalAddressBookPanel()
    End Sub
    
    Protected Sub btnDeleteFromSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DeleteFromSharedAddressBook()
        Call BindSharedAddressGrid()
    End Sub
    
    Protected Sub btnCopyToPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CopyToPersonalAddressBook()
    End Sub
    
    Protected Sub btnShowDistributionLists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowDistributionLists()
    End Sub

    Protected Sub btnCopyToPersonalAddressBookToggle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnCopyToPersonalAddressBookToggle As Button = sender
        Dim tc As TableCell
        Dim cb As CheckBox
        Dim bState As Boolean
        If btnCopyToPersonalAddressBookToggle.Text = "select all" Then
            bState = True
            btnCopyToPersonalAddressBookToggle.Text = "clear all"
        Else
            bState = False
            btnCopyToPersonalAddressBookToggle.Text = "select all"
        End If
        For Each dgi As DataGridItem In dgSharedAddressBook.Items
            tc = dgi.Cells(8)
            cb = tc.FindControl("cbCopyToPersonalAddressBook")
            cb.Checked = bState
        Next
    End Sub

    Protected Sub btnCopyToSharedAddressBookToggle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnCopyToSharedAddressBookToggle As Button = sender
        Dim tc As TableCell
        Dim cb As CheckBox
        Dim bState As Boolean
        If btnCopyToSharedAddressBookToggle.Text = "select all" Then
            bState = True
            btnCopyToSharedAddressBookToggle.Text = "clear all"
        Else
            bState = False
            btnCopyToSharedAddressBookToggle.Text = "select all"
        End If
        For Each dgi As DataGridItem In dgPersonalAddressBook.Items
            tc = dgi.Cells(8)
            cb = tc.FindControl("cbCopyToSharedAddressBook")
            cb.Checked = bState
        Next
    End Sub

    Protected Sub btnDeleteFromSharedAddressBookToggle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnDeleteFromSharedAddressBookToggle As Button = sender
        Dim tc As TableCell
        Dim cb As CheckBox
        Dim bState As Boolean
        If btnDeleteFromSharedAddressBookToggle.Text = "select all" Then
            bState = True
            btnDeleteFromSharedAddressBookToggle.Text = "clear all"
        Else
            bState = False
            btnDeleteFromSharedAddressBookToggle.Text = "select all"
        End If
        For Each dgi As DataGridItem In dgSharedAddressBook.Items
            tc = dgi.Cells(7)
            cb = tc.FindControl("cbDeleteFromSharedAddressBook")
            cb.Checked = bState
        Next
    End Sub

    Protected Sub btnDeleteFromPersonalAddressBookToggle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnDeleteFromPersonalAddressBookToggle As Button = sender
        Dim tc As TableCell
        Dim cb As CheckBox
        Dim bState As Boolean
        If btnDeleteFromPersonalAddressBookToggle.Text = "select all" Then
            bState = True
            btnDeleteFromPersonalAddressBookToggle.Text = "clear all"
        Else
            bState = False
            btnDeleteFromPersonalAddressBookToggle.Text = "select all"
        End If
        For Each dgi As DataGridItem In dgPersonalAddressBook.Items
            tc = dgi.Cells(7)
            cb = tc.FindControl("cbDeleteFromPersonalAddressBook")
            cb.Checked = bState
        Next
    End Sub

    Protected Sub dgPersonalAddressBook_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        Dim dgpcea As DataGridPageChangedEventArgs = e
        pnPage = dgpcea.NewPageIndex
        dgPersonalAddressBook.CurrentPageIndex = pnPage
        Call BindPersonalAddressGrid()
    End Sub

    Protected Sub dgSharedAddressBook_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        Dim dgpcea As DataGridPageChangedEventArgs = e
        pnPage = dgpcea.NewPageIndex
        dgSharedAddressBook.CurrentPageIndex = pnPage
        Call BindSharedAddressGrid()
    End Sub

    Protected Sub dgSharedAddressBookForDistbnLists_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        Dim dgpcea As DataGridPageChangedEventArgs = e
        pnPage = dgpcea.NewPageIndex
        dgSharedAddressBookForDistbnLists.CurrentPageIndex = pnPage
        Call BindSharedAddressGridForDistbnLists()
    End Sub

    Protected Function sNextDefaultDistributionListName() As String
        Const DEFAULT_NAME_PREFIX As String = "DistributionList"
        Dim lstDistributionListNames As List(Of String) = GetDistributionListNames()
        Dim i As Integer = 0
        Dim sNextDefaultName As String
        Do
            i += 1
            sNextDefaultName = DEFAULT_NAME_PREFIX & i.ToString
            If Not lstDistributionListNames.Contains(sNextDefaultName) Then
                Exit Do
            End If
        Loop
        sNextDefaultDistributionListName = sNextDefaultName
    End Function
    
    Protected Sub btnDistributionLists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowDistributionLists()
    End Sub
    
    Protected Sub ShowDistributionLists()
        Call GetDistributionLists()
        Call ShowDistributionListsPanel()
    End Sub
    
    Protected Function GetDistributionListNames() As List(Of String)
        Dim lstDistributionListNames As New List(Of String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT DISTINCT DistributionListName FROM AddressDistributionLists WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY DistributionListName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        While oDataReader.Read
            lstDistributionListNames.Add(oDataReader(0))
        End While
        oConn.Close()
        oConn.Dispose()
        Return lstDistributionListNames
    End Function
    
    Protected Sub GetDistributionLists()
        Dim lstDistributionListNames As List(Of String) = GetDistributionListNames()
        If lstDistributionListNames.Count > 0 Then
            'lbDistributionLists.DataTextField = lstDistributionListNames
            'lbDistributionLists.DataValueField = "DistributionListName"
            lbDistributionLists.DataSource = lstDistributionListNames
            lbDistributionLists.DataBind()
        Else
            lbDistributionLists.Items.Clear()
            lbDistributionLists.Items.Add("- no distribution lists defined -")
        End If
    End Sub
    
    Protected Sub GetDistributionList()
        Dim oConn As New SqlConnection(gsConn)
        Dim dtDistributionList As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetDistributionList", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DistributionListName", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@DistributionListName").Value = psDistributionList

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

            oAdapter.Fill(dtDistributionList)
            gvDistributionList.DataSource = dtDistributionList
            gvDistributionList.DataBind()
        Catch ex As SqlException
            ' NEED SOME ERROR HANDLING HERE
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnNewDistributionList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        psDistributionList = sNextDefaultDistributionListName()
        Call GetDistributionList()
        dgSharedAddressBookForDistbnLists.Visible = False
        Call ShowDistributionListPanel()
    End Sub
    
    Protected Sub EditDistributionList()
        If lbDistributionLists.Items.Count > 0 Then
            If lbDistributionLists.SelectedIndex < 0 Then
                WebMsgBox.Show("Please select a list to edit")
            Else
                psDistributionList = lbDistributionLists.SelectedItem.Text
                Call GetDistributionList()
                Call ShowDistributionListPanel()
            End If
        End If
    End Sub
    
    Protected Sub dgSharedAddressBookForDistbnLists_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "select" Then
            Dim itemCell As TableCell = e.Item.Cells(0)
            Call AddAddressToDistributionList(CLng(itemCell.Text))
        End If
    End Sub
    
    Protected Sub AddAddressToDistributionList(ByVal lAddressKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim dtDistributionList As New DataTable
        Dim sSQL As String
        sSQL = "SELECT CustomerKey FROM AddressDistributionLists WHERE CustomerKey = " & Session("CustomerKey") & " AND DistributionListName = '" & psDistributionList.Replace("'", "''") & "' AND GlobalAddressKey = " & lAddressKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            WebMsgBox.Show("This address has already been added to the distribution list.")
        Else
            oDataReader.Close()
            sSQL = "INSERT INTO AddressDistributionLists (GlobalAddressKey, CustomerKey, DistributionListName) VALUES (" & lAddressKey & ", " & Session("CustomerKey") & ", '" & psDistributionList.Replace("'", "''") & "')"
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
            Call GetDistributionList()
        End If
        oConn.Close()
    End Sub
    
    Protected Sub btnShowSharedAddressBookForDistbnLists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pbIsViewingGlobal = True
        tbSharedAddressBookSearchCriteriaForDistbnLists.Text = ""
        Call InitDataGrids()
        psSortExpression = "CountryName"
        Call BindSharedAddressGridForDistbnLists()
        lblInstructions.Text = String.Empty
        lblInstructions.Visible = False
    End Sub
    
    Protected Sub DeleteDistributionListEntry(ByVal DistributionListId As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "DELETE FROM AddressDistributionLists WHERE [id] = " & DistributionListId.ToString
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oCmd.ExecuteNonQuery()
        oConn.Close()
        Call GetDistributionList()
    End Sub
    
    Protected Sub lnkbtnDeleteDistbnListEntry_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call DeleteDistributionListEntry(CLng(lb.CommandArgument))
    End Sub
    
    Protected Sub lnkbtnHideAddressesForDistbnLists_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        dgSharedAddressBookForDistbnLists.Visible = False
    End Sub
    
    Protected Sub lbDistributionLists_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call EditDistributionList()
    End Sub
    
    Protected Sub btnRenameThisDistributionList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlDistributionList.Visible = False
        tbNewDistributionListName.Text = String.Empty
        tbNewDistributionListName.Focus()
        pnlRenameDistributionList.Visible = True
    End Sub
    
    Protected Sub RenameDistributionList()
        tbNewDistributionListName.Text = tbNewDistributionListName.Text.Trim
        If tbNewDistributionListName.Text = String.Empty Then
            WebMsgBox.Show("Please enter the new name for this distribution list.")
        Else
            Dim lstDistributionListNames As List(Of String) = GetDistributionListNames()
            If lstDistributionListNames.Contains(tbNewDistributionListName.Text) Then
                WebMsgBox.Show("That name is already in use.  Distribution list names must be unique. Please enter an alternative name.")
            Else
                Dim oConn As New SqlConnection(gsConn)
                Dim sSQL As String = "UPDATE AddressDistributionLists SET DistributionListName = '" & tbNewDistributionListName.Text.Replace("'", "''") & "' WHERE CustomerKey = " & Session("CustomerKey") & " AND DistributionListName = '" & psDistributionList.Replace("'", "''") & "'"
                Dim oCmd As SqlCommand
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
                oConn.Close()
                psDistributionList = tbNewDistributionListName.Text
            End If
        End If
    End Sub
    
    Protected Sub btnRename_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RenameDistributionList()
        pnlRenameDistributionList.Visible = False
        pnlDistributionList.Visible = True
    End Sub
    
    Protected Sub btnDeleteThisDistributionList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "DELETE FROM AddressDistributionLists WHERE CustomerKey = " & Session("CustomerKey") & " AND DistributionListName = '" & psDistributionList.Replace("'", "''") & "'"
        Dim oCmd As SqlCommand
        oConn.Open()
        oCmd = New SqlCommand(sSQL, oConn)
        oCmd.ExecuteNonQuery()
        oConn.Close()
        psDistributionList = Nothing
        Call ShowDistributionLists()
    End Sub

    Protected Function gvDistributionListSetRemoveVisibility()
        Return Session("EditGAB")
    End Function
    
    Protected Sub btnDefaultConsignmentDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DefaultConsignmentDestination()
    End Sub
    
    Protected Sub DefaultConsignmentDestination()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetDefaultDestination", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Dim paramAddressKey As SqlParameter = New SqlParameter("@DefaultDestinationGABKey", SqlDbType.Int, 4)
        paramAddressKey.Value = plAddressKey
        oCmd.Parameters.Add(paramAddressKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("DefaultConsignmentDestination: " & ex.ToString)
        Finally
            oConn.Close()
        End Try
        WebMsgBox.Show("From now on this will be the default address for your orders.\n\nYou can change it when you place an order.\n\n You can cancel this setting from the My Profile tab.")
    End Sub
    
    Protected Sub ddlCountry_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Call SetCountry(ddlCountry.SelectedValue, "")
    End Sub
    
    Protected Sub SetCountry(nCountryKey As Int32, sStateOrProvince As String)
        If nCountryKey = COUNTRY_CODE_USA Then
            Call SetCountryUSA(sStateOrProvince)
        ElseIf nCountryKey = COUNTRY_CODE_USA_NYC Then
            Call SetCountryUSANewYorkCity()
        ElseIf nCountryKey = COUNTRY_CODE_CANADA Then
            Call SetCountryCanada(sStateOrProvince)
        Else
            Call SetCountryOther()
        End If
    End Sub
    
    Protected Sub SetCountryOther()
        Call HideCountryRelatedControls()
        txtState.Visible = True
        lblLegendRegion.Text = "County / Region"
        lblLegendRegion.ForeColor = Drawing.Color.Blue
        txtState.Text = String.Empty
        rfvRegion.Enabled = False
    End Sub
    
    Protected Sub SetCountryUSA(sState As String)
        Call HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = True
        lblLegendRegion.Text = "State"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        Call PopulateUSStatesDropdown()
        If sState <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text = sState Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        rfvRegion.Enabled = True
    End Sub
    
    Protected Sub SetCountryUSANewYorkCity()
        Call HideCountryRelatedControls()
        lblLegendNewYorkCity.Visible = True
        lblLegendRegion.Text = "State"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        rfvRegion.Enabled = False
    End Sub
    
    Protected Sub SetCountryCanada(sProvince As String)
        Call HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = True
        lblLegendRegion.Text = "Province"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        Call PopulateCanadianProvincesDropdown()
        If sProvince <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text = sProvince Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        rfvRegion.Enabled = True
    End Sub
    
    Protected Sub HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = False
        lblLegendNewYorkCity.Visible = False
        txtState.Visible = False
    End Sub
    
    Protected Sub PopulateUSStatesDropdown()
        Dim olic As ListItemCollection = ExecuteQueryToListItemCollection("SELECT StateName + ' (' + StateAbbreviation + ')' sn, StateAbbreviation sa FROM US_States ORDER BY StateName", "sn", "sa")
        ddlUSStatesCanadianProvinces.Items.Clear()
        ddlUSStatesCanadianProvinces.Items.Add(New ListItem("- please select -", ""))
        For Each li As ListItem In olic
            ddlUSStatesCanadianProvinces.Items.Add(New ListItem(li.Text, li.Value))
        Next
    End Sub
    
    Protected Sub PopulateCanadianProvincesDropdown()
        Dim olic As ListItemCollection = ExecuteQueryToListItemCollection("SELECT ProvinceName + ' (' + ProvinceAbbreviation + ')' pn, ProvinceAbbreviation pa FROM CanadianProvinces ORDER BY ProvinceName", "pn", "pa")
        ddlUSStatesCanadianProvinces.Items.Clear()
        ddlUSStatesCanadianProvinces.Items.Add(New ListItem("- please select -", ""))
        For Each li As ListItem In olic
            ddlUSStatesCanadianProvinces.Items.Add(New ListItem(li.Text, li.Value))
        Next
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

    Property pbIsAddingAddress() As Boolean
        Get
            Dim o As Object = ViewState("IsAddingAddress")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
    
        Set(ByVal bIsAddingAddress As Boolean)
            ViewState("IsAddingAddress") = bIsAddingAddress
            ' do choreography for address detail panel; first decide whether to display add or edit panel (when editing country code cannot be changed)
            If bIsAddingAddress Then 'if adding new (user may or may not have ViewGAB privilege)
                pnlAddAddress.Visible = True
                pnlEditAddress.Visible = False
                pnlViewAddress.Visible = False
                pnlCountryDropDown.Visible = True
                pnlCountryTxtBox.Visible = False
                lblAddToPersonalAddressBook.Text = ""     ' "address added..." text
                lblAddToSharedAddressBook.Text = ""               ' "address added..." text
                If pbIsViewingGlobal Then
                    btnAddToPersonalAddressBook.Visible = False
                    btnAddToSharedAddressBook.Visible = True
                Else
                    btnAddToPersonalAddressBook.Visible = True
                    If CBool(Session("EditGAB")) Then
                        btnAddToSharedAddressBook.Visible = True
                    Else
                        btnAddToSharedAddressBook.Visible = False
                    End If
                End If
                btnAddAnotherAddress.Visible = True
                btnCloseAddressDetailPanelFromAdd.Visible = True
            Else
                pnlAddAddress.Visible = False
                pnlCountryDropDown.Visible = False
                pnlCountryTxtBox.Visible = True
                If pbIsViewingGlobal Then
                    If CBool(Session("EditGAB")) Then
                        pnlEditAddress.Visible = True
                        pnlViewAddress.Visible = False
                        btnSave.Visible = True
                        btnCloseAddressDetailPanelFromEdit.Visible = True
                    Else
                        pnlEditAddress.Visible = False
                        pnlViewAddress.Visible = True
                    End If
                Else
                    pnlEditAddress.Visible = True
                    pnlViewAddress.Visible = False
                    btnAddToSharedAddressBook.Visible = False
                    btnSave.Visible = True
                    btnCloseAddressDetailPanelFromEdit.Visible = True
                End If
            End If
        End Set
    End Property
    
    Property pbIsViewingGlobal() As Boolean
        Get
            Dim o As Object = ViewState("AB_IsViewingGlobal")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AB_IsViewingGlobal") = Value
        End Set
    End Property
    
    Property pbIsEditingDistributionList() As Boolean
        Get
            Dim o As Object = ViewState("AB_IsEditingDistributionList")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AB_IsEditingDistributionList") = Value
        End Set
    End Property
    
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("AB_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("AB_ProductOwners") = Value
        End Set
    End Property
    
    Property plAddressKey() As Long
        Get
            Dim o As Object = ViewState("AB_AddressKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("AB_AddressKey") = Value
        End Set
    End Property
    
    Property plCountryKey() As Long
        Get
            Dim o As Object = ViewState("AB_CountryKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("AB_CountryKey") = Value
        End Set
    End Property

    Property pnPage() As Integer
        Get
            Dim o As Object = ViewState("AB_Page")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("AB_Page") = Value
        End Set
    End Property
    
    Property pnVirtualItemCount() As Integer
        Get
            Dim o As Object = ViewState("AB_VirtualItemCount")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("AB_VirtualItemCount") = Value
        End Set
    End Property
    
    Property psSortExpression() As String
        Get
            Dim o As Object = ViewState("AB_SortExpression")
            If o Is Nothing Then
                Return "CountryName ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AB_SortExpression") = Value
        End Set
    End Property
    
    Property psDistributionList() As String
        Get
            Dim o As Object = ViewState("AB_DistributionList")
            If o Is Nothing Then
                Return sNextDefaultDistributionListName()
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AB_DistributionList") = Value
            lblDistributionListName.Text = Value
            lblCurrentDistributionListName.Text = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Address book</title>
</head>
<body runat="server" id="body">
    <form id="frmAddressBook" runat="server">
        <main:Header id="ctlHeader" runat="server"></main:Header>
        <table style="width: 100%" cellpadding="0" cellspacing="0"  >
            <tr class="bar_addressbook">
                <td style="width: 50%; white-space:nowrap">
                    &nbsp;<asp:Label ID="lblAddressBook" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="personal address book" />
                </td>
                <td style="width: 50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        <asp:Panel id="pnlPersonalAddressBook" runat="server" visible="False" Width="100%">
            <table width="100%">
                <tr>
                    <td style="white-space: nowrap" valign="bottom">
                        <asp:Button ID="btnShowFullAddressList" runat="server" Text="show all addresses" OnClick="btnShowFullAddressList_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class= "smallfont"  style="font-size: xx-small; font-family: Verdana;">search:</span>&nbsp;<asp:TextBox runat="server" Width="105px" Font-Size="XX-Small" Font-Names="Verdana" id="txtPersonalAddressBookSearchCriteria" MaxLength="50"></asp:TextBox>
                        <asp:Button ID="btnSearchAddress" runat="server" Text="go" OnClick="btnSearchAddress_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnPersonalAddressbookAddNewAddress" runat="server" Text="add address" OnClick="btnPersonalAddressbookAddNewAddress_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnExportPersonalAddressBook" runat="server" Text="export" OnClick="btnExportPersonalAddressBook_Click" />
                        <asp:Button ID="btnAddressUpload" runat="server" OnClientClick="window.open ('UploadToAddressBook.aspx')" Text="import" />
                    </td>
                    <td align="right" style="white-space:nowrap">
                        <asp:Button ID="btnDistributionLists1" runat="server" visible="true" Text="distribution lists" OnClick="btnShowDistributionLists_Click" />
                        <asp:Button ID="btnViewSharedAddressBook" runat="server" Text="shared address book" OnClick="btnViewSharedAddressBook_Click" />
                    </td>
                </tr>
            </table>
            <asp:DataGrid id="dgPersonalAddressBook" runat="server" Width="100%" 
                Font-Size="XX-Small" Font-Names="Arial" GridLines="None" ShowFooter="True" 
                AllowSorting="True" AutoGenerateColumns="False" 
                OnItemCommand="dgPersonalAddressBook_item_click" Visible="False" 
                OnPageIndexChanged="dgPersonalAddressBook_PageIndexChanged" PageSize="20" 
                AllowPaging="True" OnSortCommand="dgPersonalAddressBook_SortCommand">
                <HeaderStyle font-size="XX-Small" font-names="Arial" wrap="False" bordercolor="Gray"></HeaderStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle font-size="XX-Small" font-names="Arial" backcolor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="DestKey"></asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" forecolor="Blue" width="5%"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                        <HeaderTemplate>
                            Info
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:Button ID="btnViewOrEditAddress" CommandName="info" runat="server" Text="edit" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="Company" SortExpression="Company" HeaderText="Company">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Addr1" SortExpression="Addr1" HeaderText="Address Line 1">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Town" SortExpression="Town" HeaderText="City">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AttnOf" SortExpression="AttnOf" HeaderText="Attn Of">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" horizontalalign="Center" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Center"></ItemStyle>
                        <HeaderTemplate>
                            <asp:Button ID="btnDeleteFromPersonalAddressBookToggle" runat="server" Text="select all" OnClick="btnDeleteFromPersonalAddressBookToggle_Click" />
                            <br /><br />
                            delete
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="cbDeleteFromPersonalAddressBook" runat="server"></asp:CheckBox>
                        </ItemTemplate>
                        <FooterStyle wrap="False" horizontalalign="Center"></FooterStyle>
                        <FooterTemplate>
                            <asp:Button ID="btnDeleteFromPersonalAddressBook" runat="server" Text="delete" OnClick="btnDeleteFromPersonalAddressBook_Click" />
                        </FooterTemplate>
                    </asp:TemplateColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" horizontalalign="Center" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Center"></ItemStyle>
                        <HeaderTemplate>
                            <asp:Button ID="btnCopyToSharedAddressBookToggle" runat="server" Text="select all" OnClick="btnCopyToSharedAddressBookToggle_Click" />
                            <br />
                            copy to shared<br />
                            address book
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox id="cbCopyToSharedAddressBook" runat="server"></asp:CheckBox>
                        </ItemTemplate>
                        <FooterStyle wrap="False" horizontalalign="Center"></FooterStyle>
                        <FooterTemplate>
                            <asp:Button ID="btnCopyToSharedAddressBook" runat="server" Text="copy" OnClick="btnCopyToSharedAddressBook_Click" />
                        </FooterTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle HorizontalAlign="Center" PageButtonCount="20" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" Mode="NumericPages" />
            </asp:DataGrid>
            <asp:Table id="Table2" runat="server" Width="100%">
                <asp:TableRow >
                    <asp:TableCell HorizontalAlign="Left"></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right">
                        <asp:LinkButton id="btnRefreshAddressGrid" onclick="btnRefreshAddressGrid_click" runat="server" Font-Size="XX-Small" Font-Names="Arial" Visible="False" ForeColor="Blue" CausesValidation="False">refresh</asp:LinkButton>
                            &nbsp;&nbsp;&nbsp;
                        </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Label id="lblAddressMessage" runat="server" forecolor="#00C000" font-names="Arial" font-size="X-Small"></asp:Label>
        </asp:Panel>
        <asp:Panel id="pnlAddressDetail" runat="server" visible="False" Width="100%">
            <br />
            <table style="width:95%; font-size:xx-small; font-family:Verdana">
                <tr >
                    <td style="color:#0000C0; width:20%" align="right">
                        <asp:Label ID="Label1" runat="server" Font-Names="Verdana">Short Code</asp:Label>
                    </td>
                    <td style="width:30%">
                        <asp:TextBox runat="server" TabIndex="1" Width="100px" Font-Size="XX-Small" id="txtCode" MaxLength="20" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0; width:20%"></td>
                    <td style="width:30%"></td>
                </tr>
                <tr >
                    <td style="color:#0000C0" align="right">
                        &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtCompany" Font-Names="Verdana">required ></asp:RequiredFieldValidator>
                        <asp:Label ID="Label2" runat="server" ForeColor="Red" Font-Names="Verdana">Company</asp:Label></td>
                    <td>
                        <asp:TextBox runat="server" TabIndex="2" Width="150px" Font-Size="XX-Small" id="txtCompany" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0" align="right">
                        &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:RequiredFieldValidator ID="rfvCountry" ControlToValidate="ddlCountry" InitialValue="0" Text="required >" runat="server" ForeColor="Red" Font-Names="Verdana" />
                        <asp:Label ID="Label9" runat="server" ForeColor="Red" Font-Names="Verdana">Country</asp:Label>
                    </td>
                    <td >
                        <asp:Panel id="pnlCountryDropDown" runat="server" visible="False" >
                            <asp:DropDownList runat="server" TabIndex="9" Width="150px" Font-Size="XX-Small" Font-Names="Verdana" id="ddlCountry" onselectedindexchanged="ddlCountry_SelectedIndexChanged" AutoPostBack="True"/>
                        </asp:Panel>
                        <asp:Panel id="pnlCountryTxtBox" runat="server" visible="False" >
                            <asp:TextBox runat="server" TabIndex="9" Width="150px" Font-Size="XX-Small" id="txtCountry" MaxLength="50" Enabled="False" Font-Names="Verdana"/>
                        </asp:Panel>
                    </td>
                </tr>
                <tr >
                    <td style="color:#0000C0" align="right">
                        &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtAddr1" Font-Names="Verdana">required ></asp:RequiredFieldValidator>
                        <asp:Label ID="Label5" runat="server" ForeColor="Red" Font-Names="Verdana">Address Line 1</asp:Label></td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="3" Width="150px" Font-Size="XX-Small" id="txtAddr1" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0" align="right">
                        <asp:RequiredFieldValidator ID="rfvRegion" runat="server" ControlToValidate="ddlUSStatesCanadianProvinces" Enabled="false" Font-Names="Verdana">required ></asp:RequiredFieldValidator>
                        <asp:Label ID="lblLegendRegion" runat="server" Font-Names="Verdana">County / Region</asp:Label>
                    </td>
                    <td>
                        <asp:TextBox runat="server" TabIndex="7" Width="150px" Font-Size="XX-Small" id="txtState" MaxLength="50" Font-Names="Verdana"/>
                        <asp:DropDownList ID="ddlUSStatesCanadianProvinces" runat="server" Font-Size="XX-Small" Font-Names="Verdana" Visible="False"/>
                        <asp:Label ID="lblLegendNewYorkCity" runat="server" Text="NEW YORK (NY)" Visible="False" Font-Names="Verdana"/>
                    </td>
                </tr>
                <tr >
                    <td style="color:#0000C0" align="right">
                        <asp:Label ID="Label8" runat="server">Address Line 2</asp:Label>
                    </td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="4" Width="150px" Font-Size="XX-Small" id="txtAddr2" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0" align="right">
                        <asp:RequiredFieldValidator ID="rfvPostZipCode" runat="server" ControlToValidate="txtPostCode" Font-Names="Verdana">required ></asp:RequiredFieldValidator>
                        <asp:Label ID="lblLegendPostZipCode" runat="server" ForeColor="Red" Font-Names="Verdana">Post Code</asp:Label>
                    </td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="8" Width="150px" Font-Size="XX-Small" id="txtPostCode" MaxLength="10" Font-Names="Verdana"></asp:TextBox>
                    </td>
                </tr>
                <tr >
                    <td style="color:#0000C0" align="right">
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana">Address Line 3</asp:Label>
                    </td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="5" Width="150px" Font-Size="XX-Small" id="txtAddr3" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0" align="right">
                        &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana">Attn of</asp:Label></td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="10" Width="150px" Font-Size="XX-Small" id="txtAttnOf" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                </tr>
                <tr >
                    <td style="color:#0000C0" align="right">
                        &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                        <asp:RequiredFieldValidator ID="rfvTownCity" runat="server" ControlToValidate="txtCity" Font-Names="Verdana">required ></asp:RequiredFieldValidator>
                        <asp:Label ID="Label13" runat="server" ForeColor="Red" Font-Names="Verdana">Town / City</asp:Label></td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="6" Width="150px" Font-Size="XX-Small" id="txtCity" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                    <td style="color:#0000C0" align="right">
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana">Telephone</asp:Label>
                    </td>
                    <td >
                        <asp:TextBox runat="server" TabIndex="11" Width="150px" Font-Size="XX-Small" id="txtTel" MaxLength="50" Font-Names="Verdana"/>
                    </td>
                </tr>
                <tr >
                    <td colspan="4">
                        <br />
                    </td>
                </tr>
            </table>

            <asp:Panel id="pnlAddAddress" runat="server" visible="False" Width="100%">
                <table style="width: 95%; font-size: XX-Small; font-family: Verdana">
                    <tr>
                        <td style="width: 20%"></td>
                        <td style="width: 80%">
                            <asp:Button ID="btnAddToPersonalAddressBook" runat="server" Text="add to personal address book" OnClick="btnAddToPersonalAddressBook_Click" />
                            &nbsp;<asp:Button ID="btnAddToSharedAddressBook" runat="server" Text="add to shared address book" OnClick="btnAddToSharedAddressBook_Click" />
                            &nbsp;<asp:Button ID="btnAddAnotherAddress" runat="server" Text="reset" CausesValidation="false" OnClick="btnAddAnotherAddress_Click" />
                            &nbsp;<asp:Button ID="btnCloseAddressDetailPanelFromAdd" runat="server" Text="close" CausesValidation="false" OnClick="btnCloseAddressDetailPanelFromAdd_Click" />
                            &nbsp;&nbsp; <asp:Label id="lblAddToPersonalAddressBook" runat="server" forecolor="#00C000"></asp:Label>&nbsp;&nbsp; <asp:Label id="lblAddToSharedAddressBook" runat="server" forecolor="#00C000"></asp:Label>
                        </td>
                    </tr>
                </table>
            </asp:Panel>

            <asp:Panel id="pnlEditAddress" runat="server" visible="False" Width="100%">
                <table style="width: 95%; font-size: XX-Small; font-family: Verdana">
                    <tr>
                        <td style="width: 20%"></td>
                        <td style="width: 80%">
                            <asp:Button ID="btnSave" runat="server" Text="save" OnClick="btnSave_Click" />
                            &nbsp;<asp:Button ID="btnCloseAddressDetailPanelFromEdit" runat="server" Text="cancel" CausesValidation="false" OnClick="btnCloseAddressDetailPanelFromEdit_Click" />
                            &nbsp; &nbsp;
                            <asp:Button ID="btnDefaultConsignmentDestination2" runat="server" Text="always send my consignments to this destination" OnClick="btnDefaultConsignmentDestination_Click" /></td>
                    </tr>
                </table>
            </asp:Panel>

            <asp:Panel id="pnlViewAddress" runat="server" visible="False" Width="100%">
                <table style="width: 95%; font-size: XX-Small; font-family: Arial">
                    <tr>
                        <td style="width: 20%"></td>
                        <td style="width:80%">
                            <asp:Button ID="btnCloseAddressDetailPanelFromView" runat="server" Text="close" OnClick="btnCloseAddressDetailPanelFromView_Click" />
                            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:Button ID="btnDefaultConsignmentDestination1" runat="server" Text="always send my consignments to this destination" OnClick="btnDefaultConsignmentDestination_Click" /></td>
                    </tr>
                </table>
            </asp:Panel>
        </asp:Panel>

        <asp:Panel id="pnlSharedAddressBookList" runat="server" visible="False" Width="100%">
            <table style="width: 100%">
                <tr>
                    <td style="white-space: nowrap" valign="bottom">
                        <asp:Button ID="btnShowSharedAddressBook" runat="server" Text="show all addresses" OnClick="btnShowSharedAddressBook_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="smallfont" style="font-size: xx-small; font-family: Verdana;">search:</span>&nbsp;<asp:TextBox 
                            runat="server" Width="105px" Font-Size="XX-Small" Font-Names="Verdana" 
                            id="txtSharedAddressBookSearchCriteria" MaxLength="40"></asp:TextBox>
                        <asp:Button ID="btnSearchSharedAddressBook" runat="server" Text="go" OnClick="btnSearchSharedAddressBook_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnSharedAddressBookAddNewAddress" runat="server" Text="add address" OnClick="btnSharedAddressBookAddNewAddress_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnExportSharedAddressBook" runat="server" Text="export" OnClick="btnExportSharedAddressBook_Click" />
                        <asp:Button ID="btnAddressUpload2" runat="server" OnClientClick="window.open ('UploadToAddressBook.aspx')" Text="import" />
                    </td>
                    <td align="right" style="white-space:nowrap">
                        <asp:Button ID="btnDistributionLists2" runat="server" Text="distribution lists" OnClick="btnDistributionLists_Click" />
                        <asp:Button ID="btnViewPersonalAddressBook" runat="server" Text="personal address book" OnClick="btnViewPersonalAddressBook_Click" />
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td />
                </tr>
            </table>
            <asp:DataGrid id="dgSharedAddressBook" runat="server" Width="100%" 
                Font-Size="XX-Small" Font-Names="Arial" GridLines="None" ShowFooter="True" 
                AllowSorting="True" AutoGenerateColumns="False" 
                OnItemCommand="dgSharedAddressBook_item_click" Visible="False" PageSize="20" 
                OnPageIndexChanged="dgSharedAddressBook_PageIndexChanged" AllowPaging="True" 
                OnSortCommand="dgSharedAddressBook_SortCommand">
                <HeaderStyle font-size="XX-Small" font-names="Arial" wrap="False" bordercolor="Gray"></HeaderStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle font-size="XX-Small" font-names="Arial" backcolor="LightBlue"></ItemStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="DestKey"></asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" forecolor="Blue" width="5%"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                        <HeaderTemplate>
                            Info
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:Button ID="btnSharedAddressBookEdit" runat="server" CommandName="info" Text="edit" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="Company" SortExpression="Company" HeaderText="Company">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Addr1" SortExpression="Addr1" HeaderText="Address Line 1">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Town" SortExpression="Town" HeaderText="City">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AttnOf" SortExpression="AttnOf" HeaderText="Attn Of">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" horizontalalign="Center" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Center"></ItemStyle>
                        <HeaderTemplate>
                            <asp:Button ID="btnDeleteFromSharedAddressBookToggle" runat="server" Text="select all" OnClick="btnDeleteFromSharedAddressBookToggle_Click" />
                            <br /><br />
                            delete
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox id="cbDeleteFromSharedAddressBook" runat="server"></asp:CheckBox>
                        </ItemTemplate>
                        <FooterStyle wrap="False" horizontalalign="Center"></FooterStyle>
                        <FooterTemplate>
                            <asp:Button ID="btnDeleteFromSharedAddressBook" runat="server" Text="delete" OnClick="btnDeleteFromSharedAddressBook_Click" />
                        </FooterTemplate>
                    </asp:TemplateColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" horizontalalign="Center" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Center"></ItemStyle>
                        <HeaderTemplate>
                            <asp:Button ID="btnCopyToPersonalAddressBookToggle" runat="server" Text="select all" OnClick="btnCopyToPersonalAddressBookToggle_Click" />
                            <br />
                            copy to personal<br />
                            addr book
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="cbCopyToPersonalAddressBook" runat="server"></asp:CheckBox>
                        </ItemTemplate>
                        <FooterStyle wrap="False" horizontalalign="Center"></FooterStyle>
                        <FooterTemplate>
                            <asp:Button ID="btnCopyToPersonalAddressBook" runat="server" Text="copy" OnClick="btnCopyToPersonalAddressBook_Click" />
                        </FooterTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle HorizontalAlign="Center" PageButtonCount="20" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" Mode="NumericPages" />
            </asp:DataGrid>
            <asp:Table id="Table8" runat="server" Width="100%">
                <asp:TableRow >
                    <asp:TableCell HorizontalAlign="Right">
                        <asp:LinkButton id="btnRefreshSharedAddressBook" onclick="btnRefreshSharedAddressBook_click" runat="server" Font-Size="XX-Small" Font-Names="Arial" ForeColor="Blue" Visible="False" CausesValidation="False">refresh</asp:LinkButton>
                            &nbsp;&nbsp;&nbsp;
                        </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Label id="lblSharedAddressBookMessage" runat="server" forecolor="#00C000" font-names="Arial" font-size="X-Small"></asp:Label></asp:Panel>
        <asp:Panel id="pnlDistributionLists" runat="server" visible="False" Width="100%">
            <table width="100%">
                <tr>
                    <td style="white-space: nowrap" valign="bottom">
                    <asp:Label id="lblAvailableDistributionLists" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Available Distribution Lists:"/>
                    </td>
                    <td align="right" style="white-space:nowrap">
                        <asp:Button ID="btnViewSharedAddressBookFromDistbnLists" runat="server" Text="back to shared address book" OnClick="btnViewSharedAddressBook_Click" />
                    </td>
                </tr>
            </table>
            <asp:ListBox ID="lbDistributionLists" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" Rows="15" OnSelectedIndexChanged="lbDistributionLists_SelectedIndexChanged"></asp:ListBox>
            <br />
            &nbsp;<asp:Button ID="btnNewDistributionList" runat="server" Text="new list" OnClick="btnNewDistributionList_Click" />
        </asp:Panel>
        <asp:Panel id="pnlDistributionList" runat="server" visible="False" Width="100%">
            <table width="100%">
                <tr>
                    <td style="white-space: nowrap" valign="bottom">
                        <asp:Label ID="lblDistributionListName" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="Label"></asp:Label>
                            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                            <asp:Button ID="btnRenameThisDistributionList" runat="server" Text="rename this list" OnClick="btnRenameThisDistributionList_Click" />
                            <asp:Button ID="btnDeleteThisDistributionList" runat="server" Text="delete this list" OnClick="btnDeleteThisDistributionList_Click" />
                    </td>
                    <td align="right" style="white-space:nowrap">
                        <asp:Button ID="btnBackToDistributionLists" runat="server" Text="back to distribution lists" OnClick="btnDistributionLists_Click" />
                    </td>
                </tr>
            </table>
            <br />
            <asp:GridView ID="gvDistributionList" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoGenerateColumns="False" Width="100%" CellPadding="2">
                <Columns>
                    <asp:TemplateField>
                        <ItemTemplate>
                            &nbsp;<asp:LinkButton ID="lnkbtnDeleteDistbnListEntry" runat="server" CommandArgument='<%# Eval("DistributionListId") %>' Visible='<%# gvDistributionListSetRemoveVisibility() %>' Text="remove" OnClick="lnkbtnDeleteDistbnListEntry_Click" />&nbsp;
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Attn Of">
                        <ItemTemplate>
                            <asp:Label ID="lblAttnOf" runat="server" Text='<%# Eval("AttnOf") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Company">
                        <ItemTemplate>
                            <asp:Label ID="lblCompany" runat="server" Text='<%# Eval("Company") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Address Line 1">
                        <ItemTemplate>
                            <asp:Label ID="lblAddr1" runat="server" Text='<%# Eval("Addr1") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Town/City">
                        <ItemTemplate>
                            <asp:Label ID="lblTown" runat="server" Text='<%# Eval("Town") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Post Code">
                        <ItemTemplate>
                            <asp:Label ID="lblPostCode" runat="server" Text='<%# Eval("PostCode") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Country">
                        <ItemTemplate>
                            <asp:Label ID="lblCountryName" runat="server" Text='<%# Eval("CountryName") %>'/>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <EmptyDataTemplate>
                    &nbsp;- this distribution list contains no addresses -
                </EmptyDataTemplate>
                <AlternatingRowStyle BackColor="#FBFBFB" />
            </asp:GridView>
            <br />
            <asp:Button ID="btnFinishEditingList" runat="server" Text="finish editing list" OnClick="btnDistributionLists_Click" /><br />
            <br />
            <asp:Label ID="lblInstructions" runat="server" Font-Names="Verdana" ForeColor="Gray" Font-Size="XX-Small" Text="HOW TO EDIT YOUR DISTRIBUTION LIST:<br /><br />Click <b>show all available addresses</b> below (or enter a search term in the box, then click <b>go</b>) to display addresses from the address book. Click <b>select</b> to add an address to the distribution list. Click <b>remove</b> to remove an address from the distribution list. When your changes are complete, click <b>finish editing list</b>."/>
            <br />
            <table width="100%" id="tblEditDistributionList" runat="server" visible="true">
                <tr>
                    <td valign="bottom" style="height: 27px">
                        <asp:Button ID="btnShowSharedAddressBook2" runat="server" Text="show all available addresses" OnClick="btnShowSharedAddressBookForDistbnLists_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="smallfont" style="font-size: xx-small; font-family: Verdana;">search: </span> &nbsp;<asp:TextBox 
                            runat="server" Width="105px" Font-Size="XX-Small" Font-Names="Verdana" 
                            id="tbSharedAddressBookSearchCriteriaForDistbnLists" MaxLength="40"></asp:TextBox>
                        <asp:Button ID="btnSearchSharedAddressBookForDistbnLists" runat="server" Text="go" OnClick="btnSearchSharedAddressBookForDistbnLists_Click" />
                    </td>
                    <td align="right" style="white-space:nowrap; height: 27px;">
                        <asp:LinkButton ID="lnkbtnHideAddressesForDistbnLists" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnHideAddressesForDistbnLists_Click">hide addresses</asp:LinkButton>
                    </td>
                </tr>
            </table>
            <br />			
			
				<asp:DataGrid id="dgSharedAddressBookForDistbnLists" runat="server" Width="100%" Font-Size="XX-Small" 
				Font-Names="Arial" GridLines="None" ShowFooter="True" AllowSorting="True" AutoGenerateColumns="False" OnItemCommand="dgSharedAddressBookForDistbnLists_item_click" Visible="False" AllowCustomPaging="True" PageSize="20" OnPageIndexChanged="dgSharedAddressBookForDistbnLists_PageIndexChanged" 
				AllowPaging="True" OnSortCommand="dgSharedAddressBookForDistbnLists_SortCommand">
                <HeaderStyle font-size="XX-Small" font-names="Arial" wrap="False" bordercolor="Gray"></HeaderStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle font-size="XX-Small" font-names="Arial" backcolor="LightBlue"></ItemStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="DestKey"></asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle wrap="False" forecolor="Blue" width="5%"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                        <ItemTemplate>
                            <asp:Button ID="btnSharedAddressBookSelect" runat="server" CommandName="select" Text="select" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="Company" SortExpression="Company" HeaderText="Company">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Addr1" SortExpression="Addr1" HeaderText="Address Line 1">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Town" SortExpression="Town" HeaderText="City">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AttnOf" SortExpression="AttnOf" HeaderText="Attn Of">
                        <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
                <PagerStyle HorizontalAlign="Center" PageButtonCount="20" Font-Bold="False" Font-Names="Verdana" Font-Size="X-Small" Mode="NumericPages" />
            </asp:DataGrid>
        </asp:Panel>
        <asp:Panel ID="pnlRenameDistributionList" runat="server" Visible="false" Width="100%">
            <asp:Label ID="Label3x" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="Rename Distribution List"></asp:Label><br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 5%">
                    </td>
                    <td style="width: 95%">
                        <asp:Label ID="Label2x" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Rename"></asp:Label>
                        <asp:Label ID="lblCurrentDistributionListName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label>
                        <asp:Label ID="Label1x" runat="server" Text="to" Font-Names="Verdana" Font-Size="XX-Small" Height="1px"></asp:Label>
                        <asp:TextBox ID="tbNewDistributionListName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50"></asp:TextBox>
                        &nbsp;&nbsp;<asp:Button ID="btnRename" runat="server" Text="rename" OnClick="btnRename_Click" />
                    </td>
                </tr>
            </table>
        
        </asp:Panel>
        <br />
        <asp:Label id="lblError" runat="server" forecolor="Red" font-names="Arial" font-size="X-Small"></asp:Label><br />
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>