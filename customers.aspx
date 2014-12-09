<%@ Page Language="VB" Theme="AIMSDefault" %>

<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
    
    Const COLUMN_ADDRESS_START = 8
    Const COLUMN_WAREHOUSEID = 70
    Const COLUMN_EORI = 73

    Const MAX_USER_COUNT_FOR_USER_LIST As Int32 = 300
    
    Const NO_CUSTOMER_MESSAGE As String = "(no customer selected)"
    Dim gsConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call HideAllPanels()
            lblSelectedCustomer.Text = NO_CUSTOMER_MESSAGE
            psCustomerSortExpression = "CustomerAccountCode ASC"
            psUserSortExpression = "UserId ASC"
            psProductSortExpression = "ProductCode ASC"
            lblCustomersNotVisible.Visible = False
            Call PopulateWarehouseDropdown()
            Call PopulateAccountHandlerDropdown()
        End If
        tbCustomer.Focus()
        tbCustomer.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
        lblWarehouseSaved.Visible = False
        lblAccountHandlerSaved.Visible = False
    End Sub

    Protected Sub PopulateWarehouseDropdown()
        Dim sSQL As String = "SELECT WarehouseKey, WarehouseId, DeletedFlag FROM Warehouse ORDER BY WarehouseId"
        Dim dtWarehouse As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlWarehouse.Items.Clear()
        ddlWarehouse.Items.Add(New ListItem("- not defined -", 0))
        For Each drWarehouse As DataRow In dtWarehouse.Rows
            Dim sWarehouseId As String = drWarehouse("WarehouseId")
            If drWarehouse("DeletedFlag") = "Y" Then
                sWarehouseId &= " (deleted)"
            End If
            ddlWarehouse.Items.Add(New ListItem(sWarehouseId, drWarehouse("WarehouseKey")))
        Next
        ddlWarehouse.Items.Add(New ListItem("Peterborough (NO NETCOURIER)", 99))
    End Sub
    
    Protected Sub PopulateAccountHandlerDropdown()
        Dim sSQL As String = "SELECT Name, [key] 'AccountHandlerKey', DeletedFlag FROM AccountHandler ORDER BY Name"
        Dim dtAccountHandler As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlAccountHandler.Items.Clear()
        ddlAccountHandler.Items.Add(New ListItem("- not defined -", 0))
        For Each drAccountHandler As DataRow In dtAccountHandler.Rows
            Dim sName As String = drAccountHandler("Name")
            If drAccountHandler("DeletedFlag") = "1" Then
                sName &= " (deleted)"
            End If
            ddlAccountHandler.Items.Add(New ListItem(sName, drAccountHandler("AccountHandlerKey")))
        Next
    End Sub
    
    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call EnableCustomerControls()
        Call ProcessGo()
    End Sub
        
    Protected Sub EnableCustomerControls()
        btnModifyCustomerAddress.Enabled = True
        btnSaveEORI.Enabled = True
        ddlWarehouse.Enabled = True
        ddlAccountHandler.Enabled = True
        pnlCustomerAddress.Enabled = True
    End Sub
    
    Protected Sub ProcessGo()
        Call InitCustomerSearch()
        lblSelectedCustomer.Text = NO_CUSTOMER_MESSAGE
        Call SearchForCustomer()
    End Sub
    
    Protected Sub SearchForCustomer()
        tbCustomer.Text = tbCustomer.Text.Trim
        Dim sSearchString As String = tbCustomer.Text

        If Not sSearchString = String.Empty Then
            If IsNumeric(sSearchString) Then
                Call GetCustomerByKey(sSearchString)
            Else
                Call GetCustomerByName(sSearchString)
            End If
        Else
            Call GetAllCustomers()
        End If
    End Sub
    
    Protected Sub InitCustomerSearch()
        Call ResetCustomerVisibility()
        gvCustomer.Visible = True
        gvCustomer.Visible = True
        lblCustomersNotVisible.Visible = False
        psCustomerSortExpression = "CustomerAccountCode ASC"
        psUserSortExpression = "UserId ASC"
        psProductSortExpression = "ProductCode ASC"
        Call HideAllPanels()
    End Sub
    
    Protected Sub lnkbtnShowAllCustomers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        btnModifyCustomerAddress.Enabled = False
        btnSaveEORI.Enabled = False
        ddlWarehouse.Enabled = False
        ddlAccountHandler.Enabled = False
        pnlCustomerAddress.Enabled = False
        Call SelectAllCustomers()
    End Sub
    
    Protected Sub SelectAllCustomers()
        Call InitCustomerSearch()

        pnSelectedCustomerKey = -1
        Call GetAllCustomers()

        Call HideInfoPanels()
        tbCustomer.Text = String.Empty
        pnlButtons.Visible = True
        lblSelectedCustomer.Text = "<all customers>"
    End Sub
    
    Protected Sub RebindCustomerGrid()
        If lblSelectedCustomer.Text = "<all customers>" Then
            Call GetAllCustomers()
        Else
            Call SearchForCustomer()
        End If
    End Sub
    
    Protected Sub GetAllCustomers()
        'Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerStatusId = 'ACTIVE' ORDER BY " & psCustomerSortExpression
        Dim sSQL As String = GetCustomerSelectItems() & " WHERE CustomerStatusId = 'ACTIVE' ORDER BY " & psCustomerSortExpression
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvCustomer.DataSource = oDataTable
        gvCustomer.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Function GetCustomerSelectItems() As String
        GetCustomerSelectItems = "SELECT CustomerKey 'Cust Key', CustomerAccountCode 'Cust Acct Code', CustomerName 'Name', CustomerAddr1 'Addr 1', CustomerTown 'Town', CustomerPostCode 'Postcode', w.WarehouseID 'Warehouse', ah.Name 'Acct Hndlr', ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), LastJobOn, 106), ' ', '-') AS varchar(20)),'(never)') 'Last Job' FROM Customer c LEFT OUTER JOIN Warehouse w ON c.WarehouseID = w.WarehouseKey LEFT OUTER JOIN AccountHandler ah ON c.AccountHandlerKey = ah.[key] "
    End Function
    
    Protected Sub GetCustomerByKey(ByVal sSearchString As String)
        'Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & CInt(sSearchString)
        Dim sSQL As String = GetCustomerSelectItems() & " WHERE CustomerKey = " & CInt(sSearchString)
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvCustomer.DataSource = oDataTable
        gvCustomer.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub GetCustomerByName(ByVal sSearchString As String)
        'Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND CustomerAccountCode LIKE '%" & sSearchString & "%' ORDER BY " & psCustomerSortExpression
        Dim sSQL As String = GetCustomerSelectItems() & " WHERE CustomerStatusId = 'ACTIVE' AND CustomerAccountCode LIKE '%" & sSearchString & "%' ORDER BY " & psCustomerSortExpression
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvCustomer.DataSource = oDataTable
        gvCustomer.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub
    
    Protected Sub gvCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim gv As GridView = sender
        Dim nIndex As Integer = gv.SelectedIndex
        Dim gvr As GridViewRow = gv.Rows(nIndex)
        Dim nCustomerKey = gvr.Cells(1).Text
        Dim sSQL As String = "SELECT * FROM Customer WHERE CustomerKey = " & nCustomerKey
        Dim drCustomer As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        Dim sWarehouseId As String
        Dim sAccountHandlerKey As String
        
        'lblSelectedCustomer.Text = gvr.Cells(3).Text
        lblSelectedCustomer.Text = drCustomer("CustomerAccountCode")
        'tbCustomerName.Text = gvr.Cells(COLUMN_ADDRESS_START).Text
        tbCustomerName.Text = drCustomer("CustomerName")

        'tbCustomerAddress1.Text = gvr.Cells(COLUMN_ADDRESS_START + 1).Text.Replace("&nbsp;", "")
        tbCustomerAddress1.Text = (String.Empty & drCustomer("CustomerAddr1")).Replace("&nbsp;", "")

        'tbCustomerAddress2.Text = gvr.Cells(COLUMN_ADDRESS_START + 2).Text.Replace("&nbsp;", "")
        tbCustomerAddress2.Text = (String.Empty & drCustomer("CustomerAddr2")).Replace("&nbsp;", "")

        'tbCustomerAddress3.Text = gvr.Cells(COLUMN_ADDRESS_START + 3).Text.Replace("&nbsp;", "")
        tbCustomerAddress3.Text = (String.Empty & drCustomer("CustomerAddr3")).Replace("&nbsp;", "")

        'tbCustomerAddress4.Text = gvr.Cells(COLUMN_ADDRESS_START + 4).Text.Replace("&nbsp;", "")
        tbCustomerAddress4.Text = (String.Empty & drCustomer("CustomerAddr4")).Replace("&nbsp;", "")

        'tbCustomerTown.Text = gvr.Cells(COLUMN_ADDRESS_START + 5).Text.Replace("&nbsp;", "")
        tbCustomerTown.Text = (String.Empty & drCustomer("CustomerTown")).Replace("&nbsp;", "")

        'tbCustomerCounty.Text = gvr.Cells(COLUMN_ADDRESS_START + 6).Text.Replace("&nbsp;", "")
        tbCustomerCounty.Text = (String.Empty & drCustomer("CustomerCounty")).Replace("&nbsp;", "")

        'tbCustomerPostCode.Text = gvr.Cells(COLUMN_ADDRESS_START + 7).Text.Replace("&nbsp;", "")
        tbCustomerPostCode.Text = (String.Empty & drCustomer("CustomerPostCode")).Replace("&nbsp;", "")

        'tbSeparateBillingAddressFlag.Text = gvr.Cells(COLUMN_ADDRESS_START + 9).Text.Replace("&nbsp;", "")
        'tbSeparateBillingAddressFlag.Text = gvr.Cells(COLUMN_ADDRESS_START + 9).Text.Replace("&nbsp;", "")

        'tbBillingName.Text = gvr.Cells(COLUMN_ADDRESS_START + 10).Text.Replace("&nbsp;", "")
        'tbBillingName.Text = gvr.Cells(COLUMN_ADDRESS_START + 10).Text.Replace("&nbsp;", "")

        'tbBillingAddress1.Text = gvr.Cells(COLUMN_ADDRESS_START + 11).Text.Replace("&nbsp;", "")
        'tbBillingAddress1.Text = gvr.Cells(COLUMN_ADDRESS_START + 11).Text.Replace("&nbsp;", "")

        'tbBillingAddress2.Text = gvr.Cells(COLUMN_ADDRESS_START + 12).Text.Replace("&nbsp;", "")
        'tbBillingAddress2.Text = gvr.Cells(COLUMN_ADDRESS_START + 12).Text.Replace("&nbsp;", "")

        'tbBillingAddress3.Text = gvr.Cells(COLUMN_ADDRESS_START + 13).Text.Replace("&nbsp;", "")
        'tbBillingAddress4.Text = gvr.Cells(COLUMN_ADDRESS_START + 14).Text.Replace("&nbsp;", "")
        'tbBillingTown.Text = gvr.Cells(COLUMN_ADDRESS_START + 15).Text.Replace("&nbsp;", "")
        'tbBillingCounty.Text = gvr.Cells(COLUMN_ADDRESS_START + 16).Text.Replace("&nbsp;", "")
        'tbBillingPostCode.Text = gvr.Cells(COLUMN_ADDRESS_START + 17).Text.Replace("&nbsp;", "")
            
        'tbBillingAttentionOf.Text = gvr.Cells(COLUMN_ADDRESS_START + 19).Text.Replace("&nbsp;", "")
        'tbEORI.Text = gvr.Cells(COLUMN_EORI).Text.Replace("&nbsp;", "")

        'sWarehouseId = gvr.Cells(COLUMN_WAREHOUSEID).Text.Replace("&nbsp;", "")
        sWarehouseId = (String.Empty & drCustomer("WarehouseId")).Replace("&nbsp;", "")
        If sWarehouseId = String.Empty Then
            ddlWarehouse.SelectedIndex = 0
        Else
            Dim nWarehouseId As Integer = CInt(sWarehouseId)
            For i As Int32 = 1 To ddlWarehouse.Items.Count - 1
                If ddlWarehouse.Items(i).Value = sWarehouseId Then
                    ddlWarehouse.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        sAccountHandlerKey = (String.Empty & drCustomer("AccountHandlerKey")).Replace("&nbsp;", "")
        If sAccountHandlerKey = String.Empty Then
            ddlAccountHandler.SelectedIndex = 0
        Else
            Dim nAccountHandlerKey As Integer = CInt(sAccountHandlerKey)
            For i As Int32 = 1 To ddlAccountHandler.Items.Count - 1
                If ddlAccountHandler.Items(i).Value = sAccountHandlerKey Then
                    ddlAccountHandler.SelectedIndex = i
                    Exit For
                End If
            Next
        End If

        pnSelectedCustomerKey = nCustomerKey
        Call HideInfoPanels()
        pnlButtons.Visible = True
        tbEORI.Focus()
        Call EnableCustomerControls()
    End Sub
    
    Protected Sub HideAllPanels()
        pnlButtons.Visible = False
        Call HideInfoPanels()
    End Sub

    Protected Sub HideInfoPanels()
        pnlUsers.Visible = False
        pnlProducts.Visible = False
        pnlRecentWebActivity.Visible = False
        pnlCustomerAddress.Visible = False
    End Sub
    
    Protected Sub btnUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideInfoPanels()
        Dim nUserCount As Int32 = GetCountOfUsers()
        If nUserCount <= MAX_USER_COUNT_FOR_USER_LIST Then
            Call GetUsers()
            pnlUsers.Visible = True
        Else
            WebMsgBox.Show("This customer has more users (" & nUserCount.ToString & ") than can be displayed.\n\nPlease use the User Manager.")
        End If
    End Sub
    
    Protected Function GetCountOfUsers() As Int16
        GetCountOfUsers = ExecuteQueryToDataTable("SELECT COUNT (*) FROM UserProfile WHERE CustomerKey = " & pnSelectedCustomerKey).Rows(0).Item(0)
    End Function
    
    Protected Function GetUserProfileSelectItems() As String
        GetUserProfileSelectItems = "SELECT [key], UserID 'User ID', Password, FirstName 'First Name', LastName 'Last Name', Department, Type, Status, Customer 'Is Customer?', EmailAddr 'Email Addr', LastLogon 'Last Logon' FROM UserProfile WHERE "
    End Function
    
    Protected Sub GetUsers()
        Dim sSQL As String
        Dim sFilter As String = String.Empty
        If cbActiveUsers.Checked And cbSuspendedUsers.Checked Then
        Else
            If cbActiveUsers.Checked Then
                sFilter = " Status = 'Active' AND "
            End If
            If cbSuspendedUsers.Checked Then
                sFilter = " Status = 'Suspended' AND "
            End If
        End If
        If pnSelectedCustomerKey >= 0 Then
            'sSQL = GetUserProfileSelectItems() & sFilter & " CustomerKey = " & pnSelectedCustomerKey & " ORDER BY [" & psUserSortExpression & "]"
            sSQL = GetUserProfileSelectItems() & sFilter & " CustomerKey = " & pnSelectedCustomerKey & " ORDER BY " & psUserSortExpression
        Else
            'sSQL = GetUserProfileSelectItems() & sFilter & " ORDER BY [" & psUserSortExpression & "]"
            sSQL = GetUserProfileSelectItems() & sFilter & " ORDER BY " & psUserSortExpression
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        gvUsers.DataSource = oDataTable
        gvUsers.DataBind()
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Sub

    Protected Sub GetProducts()
        Dim sSQL As String
        If pnSelectedCustomerKey >= 0 Then
            sSQL = "SELECT DISTINCT lp.LogisticProductKey 'productkey', ProductCode 'Product Code', ProductDate 'Product Date', ProductDescription 'Description', ArchiveFlag 'Archived?', ProductCategory 'Category', Subcategory 'Sub-category', Misc1 'Misc 1', Misc2 'Misc 2',  Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation AS lpl ON lp.LogisticProductKey = lpl.LogisticProductKey  WHERE DeletedFlag = 'N' AND CustomerKey = " & pnSelectedCustomerKey & " ORDER BY " & psProductSortExpression
            Dim oConn As New SqlConnection(gsConn)
            Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
            Dim oDataTable As New DataTable
            oAdapter.Fill(oDataTable)
            gvProducts.DataSource = oDataTable
            gvProducts.DataBind()
            oConn.Close()
            oDataTable.Dispose()
            oAdapter.Dispose()
        Else
            WebMsgBox.Show("Cannot display products for ALL users.")
        End If
    End Sub

    Protected Sub gvProducts_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        psProductSortExpression = e.SortExpression
        Select Case e.SortExpression.ToLower
            Case "productkey"
                psProductSortExpression = "lp.LogisticProductKey" & psLogisticProductKeySortDirection
                If psLogisticProductKeySortDirection = " ASC" Then
                    psLogisticProductKeySortDirection = " DESC"
                Else
                    psLogisticProductKeySortDirection = " ASC"
                End If
            Case "product code"
                psProductSortExpression = "ProductCode" & psProductCodeSortDirection
                If psProductCodeSortDirection = " ASC" Then
                    psProductCodeSortDirection = " DESC"
                Else
                    psProductCodeSortDirection = " ASC"
                End If
            Case "product date"
                psProductSortExpression = "ProductDate" & psProductDateSortDirection
                If psProductDateSortDirection = " ASC" Then
                    psProductDateSortDirection = " DESC"
                Else
                    psProductDateSortDirection = " ASC"
                End If
            Case "description"
                psProductSortExpression = "ProductDescription" & psProductDescriptionSortDirection
                If psProductDescriptionSortDirection = " ASC" Then
                    psProductDescriptionSortDirection = " DESC"
                Else
                    psProductDescriptionSortDirection = " ASC"
                End If
            Case "archived?"
                psProductSortExpression = "ArchiveFlag" & psArchiveFlagSortDirection
                If psArchiveFlagSortDirection = " ASC" Then
                    psArchiveFlagSortDirection = " DESC"
                Else
                    psArchiveFlagSortDirection = " ASC"
                End If
            Case "category"
                psProductSortExpression = "ProductCategory" & psProductCategorySortDirection
                If psProductCategorySortDirection = " ASC" Then
                    psProductCategorySortDirection = " DESC"
                Else
                    psProductCategorySortDirection = " ASC"
                End If
            Case "sub-category"
                psProductSortExpression = "Subcategory" & psSubcategorySortDirection
                If psSubcategorySortDirection = " ASC" Then
                    psSubcategorySortDirection = " DESC"
                Else
                    psSubcategorySortDirection = " ASC"
                End If
            Case "misc 1"
                psProductSortExpression = "Misc1" & psMisc1SortDirection
                If psMisc1SortDirection = " ASC" Then
                    psMisc1SortDirection = " DESC"
                Else
                    psMisc1SortDirection = " ASC"
                End If
            Case "misc 2"
                psProductSortExpression = "Misc2" & psMisc2SortDirection
                If psMisc2SortDirection = " ASC" Then
                    psMisc2SortDirection = " DESC"
                Else
                    psMisc2SortDirection = " ASC"
                End If
            Case "quantity"
                psProductSortExpression = "Quantity" & psQuantitySortDirection
                If psQuantitySortDirection = " ASC" Then
                    psQuantitySortDirection = " DESC"
                Else
                    psQuantitySortDirection = " ASC"
                End If
                'psMisc1SortDirection
        End Select
        Call GetProducts()
    End Sub

    'Protected Sub btnRecentWebActivity_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call HideInfoPanels()
    '    Call GetRecentWebActivity()
    '    pnlRecentWebActivity.Visible = True
    'End Sub
    
    'Protected Sub GetRecentWebActivity()
    '    Dim sSQL As String
    '    If pnSelectedCustomerKey >= 0 Then
    '        sSQL = "SELECT TOP " & ddlWebHitTransactionCount.SelectedValue & " HitOn 'Hit On', ActionId Action, CustomerAccountCode Customer, up.FirstName + ' ' + up.LastName 'User' FROM LogisticWebHit lwh INNER JOIN Customer c ON lwh.CustomerKey = c.CustomerKey INNER JOIN UserProfile up ON lwh.UserKey = up.[Key] WHERE c.CustomerKey = " & pnSelectedCustomerKey & " ORDER BY HitOn DESC"
    '    Else
    '        sSQL = "SELECT TOP " & ddlWebHitTransactionCount.SelectedValue & " HitOn 'Hit On', ActionId Action, CustomerAccountCode Customer, up.FirstName + ' ' + up.LastName 'User' FROM LogisticWebHit lwh INNER JOIN Customer c ON lwh.CustomerKey = c.CustomerKey INNER JOIN UserProfile up ON lwh.UserKey = up.[Key] ORDER BY HitOn DESC"
    '    End If
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
    '    Dim oDataTable As New DataTable
    '    oAdapter.Fill(oDataTable)
    '    gvRecentWebActivity.DataSource = oDataTable
    '    gvRecentWebActivity.DataBind()
    '    oConn.Close()
    '    oDataTable.Dispose()
    '    oAdapter.Dispose()
    'End Sub

    Property psLogisticProductKeySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_LogisticProductKeySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_LogisticProductKeySortDirection") = Value
        End Set
    End Property

    Property psProductCodeSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_ProductCodeSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ProductCodeSortDirection") = Value
        End Set
    End Property

    Property psProductDateSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_ProductDateSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ProductDateSortDirection") = Value
        End Set
    End Property

    Property psProductDescriptionSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_ProductDescriptionSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ProductDescriptionSortDirection") = Value
        End Set
    End Property

    Property psArchiveFlagSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_ArchivedFlagSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ArchivedFlagSortDirection") = Value
        End Set
    End Property

    Property psProductCategorySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_ProductCategorySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ProductCategorySortDirection") = Value
        End Set
    End Property

    Property psSubcategorySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_SubcategorySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_SubcategorySortDirection") = Value
        End Set
    End Property

    Property psMisc1SortDirection() As String
        Get
            Dim o As Object = ViewState("CU_Misc1SortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_Misc1SortDirection") = Value
        End Set
    End Property

    Property psMisc2SortDirection() As String
        Get
            Dim o As Object = ViewState("CU_Misc2SortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_Misc2SortDirection") = Value
        End Set
    End Property

    Property psQuantitySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_QuantitySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_QuantitySortDirection") = Value
        End Set
    End Property

    Protected Sub gvCustomer_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        psCustomerSortExpression = e.SortExpression
        Select Case e.SortExpression.ToLower
            Case "cust key"
                psCustomerSortExpression = "CustomerKey" & psCustomerKeySortDirection
                If psCustomerKeySortDirection = " ASC" Then
                    psCustomerKeySortDirection = " DESC"
                Else
                    psCustomerKeySortDirection = " ASC"
                End If
            Case "cust acct code"
                psCustomerSortExpression = "CustomerAccountCode" & psCustomerAccountCodeSortDirection
                If psCustomerAccountCodeSortDirection = " ASC" Then
                    psCustomerAccountCodeSortDirection = " DESC"
                Else
                    psCustomerAccountCodeSortDirection = " ASC"
                End If
            Case "name"
                psCustomerSortExpression = "CustomerName"
                psCustomerSortExpression = "CustomerName" & psCustomerNameSortDirection
                If psCustomerNameSortDirection = " ASC" Then
                    psCustomerNameSortDirection = " DESC"
                Else
                    psCustomerNameSortDirection = " ASC"
                End If
            Case "addr 1"
                psCustomerSortExpression = "CustomerAddr1" & psCustomerAddr1SortDirection
                If psCustomerAddr1SortDirection = " ASC" Then
                    psCustomerAddr1SortDirection = " DESC"
                Else
                    psCustomerAddr1SortDirection = " ASC"
                End If
            Case "town"
                psCustomerSortExpression = "CustomerTown" & psCustomerTownSortDirection
                If psCustomerTownSortDirection = " ASC" Then
                    psCustomerTownSortDirection = " DESC"
                Else
                    psCustomerTownSortDirection = " ASC"
                End If
            Case "postcode"
                psCustomerSortExpression = "CustomerPostcode" & psCustomerPostcodeSortDirection
                If psCustomerPostcodeSortDirection = " ASC" Then
                    psCustomerPostcodeSortDirection = " DESC"
                Else
                    psCustomerPostcodeSortDirection = " ASC"
                End If
            Case "warehouse"
                psCustomerSortExpression = "c.WarehouseId" & psCustomerWarehouseSortDirection
                If psCustomerWarehouseSortDirection = " ASC" Then
                    psCustomerWarehouseSortDirection = " DESC"
                Else
                    psCustomerWarehouseSortDirection = " ASC"
                End If
            Case "acct hndlr"
                psCustomerSortExpression = "AccountHandlerKey" & psCustomerAccountHandlerSortDirection
                If psCustomerAccountHandlerSortDirection = " ASC" Then
                    psCustomerAccountHandlerSortDirection = " DESC"
                Else
                    psCustomerAccountHandlerSortDirection = " ASC"
                End If
            Case "last job"
                psCustomerSortExpression = "LastJobOn" & psCustomerLastJobOnSortDirection
                If psCustomerLastJobOnSortDirection = " ASC" Then
                    psCustomerLastJobOnSortDirection = " DESC"
                Else
                    psCustomerLastJobOnSortDirection = " ASC"
                End If
        End Select
        If tbCustomer.Text = String.Empty Then
            Call GetAllCustomers()
        Else
            Call SearchForCustomer()
        End If
    End Sub
    
    Protected Sub gvUsers_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        psUserSortExpression = e.SortExpression
      
        Select Case e.SortExpression.ToLower
            Case "key"
                psUserSortExpression = "[key]" & psUserKeySortDirection
                If psUserKeySortDirection = " ASC" Then
                    psUserKeySortDirection = " DESC"
                Else
                    psUserKeySortDirection = " ASC"
                End If
            Case "user id"
                psUserSortExpression = "UserID" & psUserIDSortDirection
                If psUserIDSortDirection = " ASC" Then
                    psUserIDSortDirection = " DESC"
                Else
                    psUserIDSortDirection = " ASC"
                End If
            Case "password"
                psUserSortExpression = "Password" & psPasswordSortDirection
                If psPasswordSortDirection = " ASC" Then
                    psPasswordSortDirection = " DESC"
                Else
                    psPasswordSortDirection = " ASC"
                End If
            Case "first name"
                psUserSortExpression = "FirstName" & psFirstNameSortDirection
                If psFirstNameSortDirection = " ASC" Then
                    psFirstNameSortDirection = " DESC"
                Else
                    psFirstNameSortDirection = " ASC"
                End If
            Case "last name"
                psUserSortExpression = "LastName" & psLastNameSortDirection
                If psLastNameSortDirection = " ASC" Then
                    psLastNameSortDirection = " DESC"
                Else
                    psLastNameSortDirection = " ASC"
                End If
            Case "department"
                psUserSortExpression = "Department" & psDepartmentSortDirection
                If psDepartmentSortDirection = " ASC" Then
                    psDepartmentSortDirection = " DESC"
                Else
                    psDepartmentSortDirection = " ASC"
                End If
            Case "type"
                psUserSortExpression = "Type" & psUserTypeSortDirection
                If psUserTypeSortDirection = " ASC" Then
                    psUserTypeSortDirection = " DESC"
                Else
                    psUserTypeSortDirection = " ASC"
                End If
            Case "status"
                psUserSortExpression = "Status" & psUserStatusSortDirection
                If psUserStatusSortDirection = " ASC" Then
                    psUserStatusSortDirection = " DESC"
                Else
                    psUserStatusSortDirection = " ASC"
                End If
            Case "is customer?"
                psUserSortExpression = "Customer" & psIsCustomerSortDirection
                If psIsCustomerSortDirection = " ASC" Then
                    psIsCustomerSortDirection = " DESC"
                Else
                    psIsCustomerSortDirection = " ASC"
                End If
            Case "email addr"
                psUserSortExpression = "EmailAddr" & psEmailAddrSortDirection
                If psEmailAddrSortDirection = " ASC" Then
                    psEmailAddrSortDirection = " DESC"
                Else
                    psEmailAddrSortDirection = " ASC"
                End If
            Case "last logon"
                psUserSortExpression = "LastLogon" & psLastLogonSortDirection
                If psLastLogonSortDirection = " ASC" Then
                    psLastLogonSortDirection = " DESC"
                Else
                    psLastLogonSortDirection = " ASC"
                End If
        End Select
        Call GetUsers()
    End Sub

    Protected Sub ResetCustomerVisibility()
        lnkbtnCustomerVisibility.Text = "hide customers"
    End Sub
    
    Protected Sub lnkbtnCustomerVisibility_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If lnkbtnCustomerVisibility.Text = "hide customers" Then
            lnkbtnCustomerVisibility.Text = "show customers"
        Else
            lnkbtnCustomerVisibility.Text = "hide customers"
        End If
        gvCustomer.Visible = Not gvCustomer.Visible
        lblCustomersNotVisible.Visible = Not lblCustomersNotVisible.Visible
    End Sub
    
    Protected Sub btnModifyCustomerAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' Call HideInfoPanels()
        If pnlCustomerAddress.Visible = False Then
            pnlCustomerAddress.Visible = True
            btnModifyCustomerAddress.Text = "hide customer address"
        Else
            pnlCustomerAddress.Visible = False
            btnModifyCustomerAddress.Text = "modify customer address"
        End If
    End Sub
    
    Protected Sub btnUpdateAddressDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("UPDATE Customer SET CustomerName = '")
        sbSQL.Append(tbCustomerName.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerAddr1 = '")
        sbSQL.Append(tbCustomerAddress1.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerAddr2 = '")
        sbSQL.Append(tbCustomerAddress2.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerAddr3 = '")
        sbSQL.Append(tbCustomerAddress3.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerAddr4 = '")
        sbSQL.Append(tbCustomerAddress4.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerTown = '")
        sbSQL.Append(tbCustomerTown.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerCounty = '")
        sbSQL.Append(tbCustomerCounty.Text.Replace("'", "''"))
        sbSQL.Append("', ")
        sbSQL.Append("CustomerPostCode = '")
        sbSQL.Append(tbCustomerPostCode.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("SeparateBillingAddressFlag = '")
        'sbSQL.Append(tbSeparateBillingAddressFlag.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingName = '")
        'sbSQL.Append(tbBillingName.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingAddr1 = '")
        'sbSQL.Append(tbBillingAddress1.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingAddr2 = '")
        'sbSQL.Append(tbBillingAddress2.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingAddr3 = '")
        'sbSQL.Append(tbBillingAddress3.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingAddr4 = '")
        'sbSQL.Append(tbBillingAddress4.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingTown = '")
        'sbSQL.Append(tbBillingTown.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingCounty = '")
        'sbSQL.Append(tbBillingCounty.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingPostCode = '")
        'sbSQL.Append(tbBillingPostCode.Text.Replace("'", "''"))
        'sbSQL.Append("', ")
        'sbSQL.Append("BillingAttentionOf = '")

        'sbSQL.Append(tbBillingAttentionOf.Text.Replace("'", "''"))
        sbSQL.Append("' WHERE CustomerKey = ")
        sbSQL.Append(pnSelectedCustomerKey)
        Dim sSQL As String = sbSQL.ToString
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oCmd.ExecuteNonQuery()
        oConn.Close()
        Call RebindCustomerGrid()
    End Sub
    
    Protected Sub ddlWarehouse_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL As New StringBuilder
        If ddl.SelectedItem.Text.ToLower.Contains("(deleted)") Then
            WebMsgBox.Show("You cannot set the warehouse to a deleted warehouse entry.\n\nPlease select a different warehouse.")
            Dim nOriginalWarehouse As Int32 = ExecuteQueryToDataTable("SELECT ISNULL(WarehouseID, 0) FROM Customer WHERE CustomerKey = " & pnSelectedCustomerKey).Rows(0).Item(0)
            For i As Int32 = 0 To ddl.Items.Count - 1
                If ddl.Items(i).Value = nOriginalWarehouse Then
                    ddl.SelectedIndex = i
                    Exit Sub
                End If
            Next
        End If
        sbSQL.Append("UPDATE Customer SET WarehouseId = ")
        If ddl.SelectedIndex = 0 Then
            sbSQL.Append("NULL")
        Else
            sbSQL.Append(ddl.SelectedValue)
        End If
        sbSQL.Append(" WHERE CustomerKey = ")
        sbSQL.Append(pnSelectedCustomerKey)
        Dim sSQL As String = sbSQL.ToString
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oCmd.ExecuteNonQuery()
        oConn.Close()
        lblWarehouseSaved.Visible = True
        Call RebindCustomerGrid()
    End Sub
    
    Protected Sub gvUsers_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Const ROW_TEMPLATE As Int32 = 0
        Const ROW_PASSWORD As Int32 = 3
        Dim gvrea As GridViewRowEventArgs = e
        Dim oPassword As SprintInternational.Password = New SprintInternational.Password()
        If gvrea.Row.RowType = DataControlRowType.DataRow Then
            gvrea.Row.Cells(ROW_PASSWORD).Text = oPassword.Decrypt(gvrea.Row.Cells(ROW_PASSWORD).Text)
            Dim hidStatus As HiddenField = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("hidStatus")
            Dim lnkbtnActivateUser As LinkButton = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("lnkbtnActivateUser")
            Dim lnkbtnSuspendUser As LinkButton = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("lnkbtnSuspendUser")
            If hidStatus.Value.ToLower = "active" Then
                lnkbtnSuspendUser.Visible = True
                lnkbtnActivateUser.Visible = False
            Else
                lnkbtnSuspendUser.Visible = False
                lnkbtnActivateUser.Visible = True
                gvrea.Row.BackColor = Drawing.Color.Silver
            End If
        End If
    End Sub

    Protected Sub gvProducts_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Const ROW_TEMPLATE As Int32 = 0
        Dim gvrea As GridViewRowEventArgs = e
        If gvrea.Row.RowType = DataControlRowType.DataRow Then
            Dim hidArchiveFlag As HiddenField = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("hidArchiveFlag")
            Dim lnkbtnArchiveProduct As LinkButton = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("lnkbtnArchiveProduct")
            Dim lnkbtnUnarchiveProduct As LinkButton = gvrea.Row.Cells(ROW_TEMPLATE).FindControl("lnkbtnUnarchiveProduct")
            If hidArchiveFlag.Value.ToLower = "n" Then
                lnkbtnArchiveProduct.Visible = True
                lnkbtnUnarchiveProduct.Visible = False
            Else
                lnkbtnArchiveProduct.Visible = False
                lnkbtnUnarchiveProduct.Visible = True
                gvrea.Row.BackColor = Drawing.Color.Silver
            End If
        End If
    End Sub

    Protected Sub lnkbtnArchiveProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteNonQuery("UPDATE LogisticProduct SET ArchiveFlag = 'Y' WHERE LogisticProductKey = " & lnkbtn.CommandArgument)
        Call GetProducts()
    End Sub

    Protected Sub lnkbtnUnarchiveProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteNonQuery("UPDATE LogisticProduct SET ArchiveFlag = 'N' WHERE LogisticProductKey = " & lnkbtn.CommandArgument)
        Call GetProducts()
    End Sub

    Protected Sub lnkbtnActivateUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteNonQuery("UPDATE UserProfile SET Status = 'Active' WHERE [key] = " & lnkbtn.CommandArgument)
        Call GetUsers()
    End Sub

    Protected Sub lnkbtnSuspendUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteNonQuery("UPDATE UserProfile SET Status = 'Suspended' WHERE [key] = " & lnkbtn.CommandArgument)
        Call GetUsers()
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

    Protected Sub btnSaveEORI_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not ExecuteNonQuery("UPDATE Customer SET EORI = '" & tbEORI.Text & "' WHERE CustomerKey = " & pnSelectedCustomerKey) Then
            WebMsgBox.Show("Error saving EORI")
        Else
            Call SearchForCustomer()
        End If
    End Sub

    Property pnSelectedCustomerKey() As Integer
        Get
            Dim o As Object = ViewState("CU_SelectedCustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CU_SelectedCustomerKey") = Value
        End Set
    End Property
    
    Property psCustomerSortExpression() As String
        Get
            Dim o As Object = ViewState("CU_CustomerSortExpression")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerSortExpression") = Value
        End Set
    End Property
    
    Property psUserSortExpression() As String
        Get
            Dim o As Object = ViewState("CU_UserSortExpression")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_UserSortExpression") = Value
        End Set
    End Property

    Property psProductSortExpression() As String
        Get
            Dim o As Object = ViewState("CU_ProductSortExpression")
            If o Is Nothing Then
                Return -1
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_ProductSortExpression") = Value
        End Set
    End Property

    Property psCustomerKeySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerKeySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerKeySortDirection") = Value
        End Set
    End Property

    Property psCustomerAccountCodeSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerAccountCodeSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerAccountCodeSortDirection") = Value
        End Set
    End Property

    Property psCustomerNameSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerNameSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerNameSortDirection") = Value
        End Set
    End Property

    Property psCustomerAddr1SortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerAddr1SortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerAddr1SortDirection") = Value
        End Set
    End Property

    Property psCustomerTownSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerTownSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerTownSortDirection") = Value
        End Set
    End Property

    Property psCustomerPostcodeSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerPostcodeSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerPostcodeSortDirection") = Value
        End Set
    End Property

    Property psCustomerWarehouseSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerWarehouseSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerWarehouseSortDirection") = Value
        End Set
    End Property

    Property psCustomerAccountHandlerSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerAccountHandlerSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerAccountHandlerSortDirection") = Value
        End Set
    End Property

    Property psCustomerLastJobOnSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_CustomerLastJobOnSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_CustomerLastJobOnSortDirection") = Value
        End Set
    End Property

    Property psUserKeySortDirection() As String
        Get
            Dim o As Object = ViewState("CU_UserKeySortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_UserKeySortDirection") = Value
        End Set
    End Property

    Property psUserIDSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_UserIDSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_UserIDSortDirection") = Value
        End Set
    End Property

    Property psPasswordSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_PasswordSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_PasswordSortDirection") = Value
        End Set
    End Property

    Property psFirstNameSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_FIrstNameSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_FIrstNameSortDirection") = Value
        End Set
    End Property

    Property psLastNameSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_LastNameSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_LastNameSortDirection") = Value
        End Set
    End Property

    Property psDepartmentSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_DepartmentSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_DepartmentSortDirection") = Value
        End Set
    End Property

    Property psUserTypeSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_UserTypeSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_UserTypeSortDirection") = Value
        End Set
    End Property

    Property psUserStatusSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_UserStatusSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_UserStatusSortDirection") = Value
        End Set
    End Property

    Property psIsCustomerSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_IsCustomerSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_IsCustomerSortDirection") = Value
        End Set
    End Property

    Property psEmailAddrSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_EmailAddrSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_EmailAddrSortDirection") = Value
        End Set
    End Property

    Property psLastLogonSortDirection() As String
        Get
            Dim o As Object = ViewState("CU_LastLogonSortDirection")
            If o Is Nothing Then
                Return " ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CU_LastLogonSortDirection") = Value
        End Set
    End Property

    Protected Sub cbActiveUsers_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked = False Then
            cbSuspendedUsers.Checked = True
        End If
        Call GetUsers()
    End Sub

    Protected Sub cbSuspendedUsers_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked = False Then
            cbActiveUsers.Checked = True
        End If
        Call GetUsers()
    End Sub
    
    Protected Sub btnProducts_Click(sender As Object, e As System.EventArgs)
        Call HideInfoPanels()
        Call GetProducts()
        pnlProducts.Visible = True
    End Sub
    
    Protected Sub ddlAccountHandler_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL As New StringBuilder
        If ddl.SelectedItem.Text.ToLower.Contains("(deleted)") Then
            WebMsgBox.Show("You cannot set the warehouse to a deleted account handler entry.\n\nPlease select a different account handler.")
            Dim nOriginalAccountHandler As Int32 = ExecuteQueryToDataTable("SELECT ISNULL(AccountHandlerKey, 0) FROM Customer WHERE CustomerKey = " & pnSelectedCustomerKey).Rows(0).Item(0)
            For i As Int32 = 0 To ddl.Items.Count - 1
                If ddl.Items(i).Value = nOriginalAccountHandler Then
                    ddl.SelectedIndex = i
                    Exit Sub
                End If
            Next
        End If
        sbSQL.Append("UPDATE Customer SET AccountHandlerKey = ")
        If ddl.SelectedIndex = 0 Then
            sbSQL.Append("NULL")
        Else
            sbSQL.Append(ddl.SelectedValue)
        End If
        sbSQL.Append(" WHERE CustomerKey = ")
        sbSQL.Append(pnSelectedCustomerKey)
        Dim sSQL As String = sbSQL.ToString
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oCmd.ExecuteNonQuery()
        oConn.Close()
        lblAccountHandlerSaved.Visible = True
        Call RebindCustomerGrid()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml ">
<head runat="server">
    <title>Customer Info</title>
    <link href="Reports.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
        <strong>Customer Info<br />
        </strong>
        <br />
        <table style="width: 100%">
            <tr>
                <td style="width: 35%" align="right">
                </td>
                <td style="width: 65%" align="left">
                </td>
            </tr>
            <tr>
                <td>
                    Search for full or partial <strong>Customer Name</strong>, or CustomerKey value:
                </td>
                <td>
                    <asp:TextBox ID="tbCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                    <asp:Button ID="btnGo" runat="server" Text="go" OnClick="btnGo_Click" />
                    &nbsp; &nbsp; &nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnShowAllCustomers" runat="server" OnClick="lnkbtnShowAllCustomers_Click">show all customers</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="lblBidrectionalSortReminder01" runat="server" Font-Bold="True" 
            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#FF6600" Text="Click column headings for bi-directional sorting"/>
                </td>
            </tr>
        </table>
    <br />
    <asp:GridView ID="gvCustomer" runat="server" Width="100%" AutoGenerateSelectButton="True"
        CellPadding="2" OnSelectedIndexChanged="gvCustomer_SelectedIndexChanged" AllowSorting="True"
        OnSorting="gvCustomer_Sorting" Font-Names="Verdana" Font-Size="XX-Small">
    </asp:GridView>
    <span style="color: #ff0000">
        <asp:Label ID="lblCustomersNotVisible" runat="server" ForeColor="Red" Text=" customers not current visible - click show customers to see customer list"></asp:Label></span><br />
    <br />
    Selected customer: &nbsp;<asp:Label ID="lblSelectedCustomer" runat="server" Font-Bold="True"></asp:Label><br />
    <br />
    <asp:Panel ID="pnlButtons" runat="server" Width="100%">
        <asp:Button ID="btnUsers" runat="server" Text="show users" OnClick="btnUsers_Click" />&nbsp;<asp:CheckBox
            ID="cbActiveUsers" runat="server" Checked="True" Text="active" AutoPostBack="True"
            OnCheckedChanged="cbActiveUsers_CheckedChanged" />&nbsp;<asp:CheckBox ID="cbSuspendedUsers"
                runat="server" Text="suspended" AutoPostBack="True" OnCheckedChanged="cbSuspendedUsers_CheckedChanged" />
        &nbsp;&nbsp;&nbsp; &nbsp;
        <asp:Button ID="btnModifyCustomerAddress" runat="server" 
            OnClick="btnModifyCustomerAddress_Click" Text="modify customer address..." />
        <%--&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp;
        <asp:Button ID="btnRecentWebActivity" runat="server" Text="recent web activity" OnClick="btnRecentWebActivity_Click" />
        last
        <asp:DropDownList ID="ddlWebHitTransactionCount" runat="server" Font-Names="Verdana"
            Font-Size="XX-Small">
            <asp:ListItem>10</asp:ListItem>
            <asp:ListItem>50</asp:ListItem>
            <asp:ListItem>200</asp:ListItem>
        </asp:DropDownList>
        txactions &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;--%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:LinkButton ID="lnkbtnCustomerVisibility" runat="server" 
            OnClick="lnkbtnCustomerVisibility_Click">hide customers</asp:LinkButton>
        <br />
        EORI:
        <asp:TextBox ID="tbEORI" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" MaxLength="21" Width="100px"></asp:TextBox>
        &nbsp;<asp:Button ID="btnSaveEORI" runat="server" OnClick="btnSaveEORI_Click" 
            Text="save EORI" />
        &nbsp; &nbsp; &nbsp; &nbsp;Warehouse:&nbsp;
        <asp:DropDownList ID="ddlWarehouse" runat="server" 
            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" 
            OnSelectedIndexChanged="ddlWarehouse_SelectedIndexChanged">
        </asp:DropDownList>
        &nbsp;<asp:Label ID="lblWarehouseSaved" runat="server" Font-Bold="True" 
            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#33CC33" Text="saved" 
            Visible="False"></asp:Label>
        &nbsp; &nbsp; &nbsp; &nbsp; Acct Hndlr:
        <asp:DropDownList ID="ddlAccountHandler" runat="server" 
            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" 
            onselectedindexchanged="ddlAccountHandler_SelectedIndexChanged"/>        &nbsp;<asp:Label 
            ID="lblAccountHandlerSaved" runat="server" Font-Bold="True" 
            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#33CC33" Text="saved" 
            Visible="False"></asp:Label>
        <br />
        <br />
        <asp:Button ID="btnProducts" runat="server" onclick="btnProducts_Click" 
            Text="show products" />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlCustomerAddress" runat="server" Width="100%">
        <strong>
            <br />
            Customer &amp; Billing Address</strong><br />
        <br />
        <table style="width: 100%; font-size: x-small; font-family: Verdana;">
            <tr>
                <td style="width: 25%; height: 22px;" align="right">
                    Customer Name
                </td>
                <td style="width: 2%; height: 22px;">
                </td>
                <td style="width: 73%; height: 22px;">
                    <asp:TextBox ID="tbCustomerName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Address 1
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Address 2
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Address 3
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerAddress3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Address 4
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerAddress4" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Town
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer County
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerCounty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    Customer Post Code
                </td>
                <td>
                </td>
                <td>
                    <asp:TextBox ID="tbCustomerPostCode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnUpdateAddressDetails" runat="server" 
                        OnClick="btnUpdateAddressDetails_Click" Text="update address details" />
                </td>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <%--<tr>
                    <td align="right">
                        Separate Billing Address Flag (Y/N)</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbSeparateBillingAddressFlag" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Name</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Address 1</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Address 2</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="height: 14px">
                        Billing Address 3</td>
                    <td style="height: 14px">
                    </td>
                    <td style="height: 14px">
                        <asp:TextBox ID="tbBillingAddress3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Address 4</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingAddress4" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Town</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing County</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingCounty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Post Code</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingPostCode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right">
                        Billing Attention Of</td>
                    <td>
                    </td>
                    <td>
                        <asp:TextBox ID="tbBillingAttentionOf" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px"></asp:TextBox></td>
                </tr>--%>
        </table>
        <br />
        &nbsp;</asp:Panel>
    <br />
    <asp:Panel ID="pnlUsers" runat="server" Width="100%">
        <b>Users</b>&nbsp;&nbsp;
        <asp:Label ID="lblBidrectionalSortReminder2" runat="server" Font-Bold="True" 
            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#FF6600" 
            Text="Click column headings for bi-directional sorting"/>
        <br />
        <br />
        <asp:GridView ID="gvUsers" runat="server" AllowSorting="True" CellPadding="2" 
            Font-Names="Verdana" Font-Size="XX-Small" OnRowDataBound="gvUsers_RowDataBound" 
            OnSorting="gvUsers_Sorting" Width="100%">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnActivateUser" runat="server" 
                            CommandArgument='<%# Bind("key") %>' Font-Names="Verdana" Font-Size="XX-Small" 
                            OnClick="lnkbtnActivateUser_Click">activate</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtnSuspendUser" runat="server" 
                            CommandArgument='<%# Bind("key") %>' Font-Names="Verdana" Font-Size="XX-Small" 
                            OnClick="lnkbtnSuspendUser_Click">suspend</asp:LinkButton>
                        <asp:HiddenField ID="hidStatus" runat="server" Value='<%# Bind("Status") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </asp:Panel>
    <br />
    <br />
    <asp:Panel ID="pnlProducts" runat="server" Width="100%">
        <b>Products</b>&nbsp;&nbsp;
        <asp:Label ID="lblBidrectionalSortReminder3" runat="server" Font-Bold="True" 
            Font-Names="Verdana" Font-Size="XX-Small" ForeColor="#FF6600" 
            Text="Click column headings for bi-directional sorting"/>
        <br />
        <br />
        <asp:GridView ID="gvProducts" runat="server" AllowSorting="True" 
            CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" 
            OnRowDataBound="gvProducts_RowDataBound" OnSorting="gvProducts_Sorting" 
            Width="100%">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnUnarchiveProduct" runat="server" 
                            CommandArgument='<%# Bind("ProductKey") %>' Font-Names="Verdana" 
                            Font-Size="XX-Small" OnClick="lnkbtnUnarchiveProduct_Click">unarchive</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtnArchiveProduct" runat="server" 
                            CommandArgument='<%# Bind("ProductKey") %>' Font-Names="Verdana" 
                            Font-Size="XX-Small" OnClick="lnkbtnArchiveProduct_Click">archive</asp:LinkButton>
                        <asp:HiddenField ID="hidArchiveFlag" runat="server" 
                            Value='<%# Bind("[Archived?]") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </asp:Panel>
    <br />
    <br />
    <asp:Panel ID="pnlRecentWebActivity" runat="server" Width="100%">
        <strong>Recent Web Activity</strong><br />
        <br />
        <br />
        <asp:GridView ID="gvRecentWebActivity" runat="server" CellPadding="2" Width="100%"
            Font-Names="Verdana" Font-Size="XX-Small">
        </asp:GridView>
    </asp:Panel>
    </form>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
