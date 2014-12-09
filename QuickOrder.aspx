<%@ Page Language="VB" Theme="QuickOrder" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    'TEST!!!
    'support multiple destinations
    'add service level
    'remove CRLF from Special Instrs & Packing Note fields
    
    Const ITEMS_PER_REQUEST As Integer = 30
    Const COUNTRY_UK As Int32 = 222
    Const ACCOUNT_CODE As String = "COURI11111"
    Const LICENSE_KEY As String = "RA61-XZ94-CT55-FH67"

    Const COUNTRY_CODE_CANADA As Int32 = 38
    Const COUNTRY_CODE_USA As Int32 = 223
    Const COUNTRY_CODE_USA_NYC As Int32 = 256
    
    Const CUSTOMER_INTERNAL As Int32 = 566
    Const CUSTOMER_ARTHRITIS As Int32 = 711

    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gdtBasket As DataTable
   
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            If CInt(Session("CustomerKey")) = CUSTOMER_INTERNAL Then
                Call GetCustomerAccountCodes()
                Call SetCustomerOptions()
                divInternal.Visible = True
                divMainForm.Visible = False
                ddlCustomer.Focus()
            Else
                pnImpersonateCustomer = CInt(Session("CustomerKey"))
                pnImpersonateBookedByUser = CInt(Session("UserKey"))
                divMainForm.Visible = True
                Call GetSiteFeatures()
                'PopulateProductDropdown(pnImpersonateCustomer)	   ' XXXX
                'ddlProduct.Focus()
                rcbProduct.Focus()
            End If
            Call GetCountries()
            Session("BO_BasketData") = Nothing
            SetAddressVisibility(False)
            psVirtualThumbURL = ConfigLib.GetConfigItem_Virtual_Thumb_URL
            'psVirtualThumbURL("http://my.transworld.eu.com/common/")
            'tbSearch.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchGo.ClientID + "')")
            'Call SetFilterControlsVisibility(False)
            Call ShowEmptyBasket()
            Call SetCustomerOptions()
        End If
    End Sub
   
    Protected Sub SetCustomerOptions()
        If IsArthritis() Then
            Call HideDefaultCustRefFields()
            trArthritisCostCentres.Visible = True
            trArthritisCategories.Visible = True
            trArthritisPONumber.Visible = True
            trArthritisAdvisory.Visible = True
        End If
    End Sub
    
    Protected Sub HideDefaultCustRefFields()
        trCustRef1.Visible = False
        trCustRef2.Visible = False
        trCustRef3.Visible = False
        trCustRef4.Visible = False
    End Sub
    
    Protected Function IsArthritis() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsArthritis = IIf(gbSiteTypeDefined, gsSiteType = "arthritis", nCustomerKey = CUSTOMER_ARTHRITIS)
    End Function
    
    Protected Sub ShowEmptyBasket()
        Dim dt As DataTable = Nothing
        gvBasket.DataSource = dt
        gvBasket.DataBind()
        Call SetPlaceOrderButtonVisibility()
    End Sub
    
    Protected Sub GetSiteFeatures()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent2", oConn)
        
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
        'pbSiteShowZeroStockBalances = dr("ShowZeroStock")
        'pbSiteApplyMaxGrabs = dr("ApplyMaxGrabs")
        'pbMultipleAddressOrders = dr("MultipleAddressOrders")
        'pbOrderAuthorisation = dr("OrderAuthorisation")
        'pbProductAuthorisation = dr("ProductAuthorisation")
        'pbCalendarManagement = dr("CalendarManagement")
        
        'pbCustomLetters = dr("CustomLetters")
        'pbCustomLetters = False
        
        'pbOnDemandProducts = dr("OnDemandProducts")
        'pbZeroStockNotifications = dr("Misc3")
        'pbShowNotes = dr("ShowNotes")
        'pnCategoryMode = dr("CategoryCount")
        
        If CBool(dr("StockOrderCustRef1Visible")) Then
            trCustRef1.Visible = True
            lblLegendCustRef1.Text = dr("StockOrderCustRefLabel1Legend") & ":"
            If CBool(dr("StockOrderCustRef1Mandatory")) Then
                lblLegendCustRef1.ForeColor = Drawing.Color.Red
                rfvCustRef1.Enabled = True
                rfvCustRef1.EnableClientScript = True
            Else
                rfvCustRef1.Enabled = False
                rfvCustRef1.EnableClientScript = False
            End If
        Else
            trCustRef1.Visible = False
        End If
        If CBool(dr("StockOrderCustRef2Visible")) Then
            trCustRef2.Visible = True
            lblLegendCustRef2.Text = dr("StockOrderCustRefLabel2Legend") & ":"
            If CBool(dr("StockOrderCustRef2Mandatory")) Then
                lblLegendCustRef2.ForeColor = Drawing.Color.Red
                rfvCustRef2.Enabled = True
                rfvCustRef2.EnableClientScript = True
            Else
                rfvCustRef2.Enabled = False
                rfvCustRef2.EnableClientScript = False
            End If
        Else
            trCustRef2.Visible = False
        End If
        If CBool(dr("StockOrderCustRef3Visible")) Then
            trCustRef3.Visible = True
            lblLegendCustRef3.Text = dr("StockOrderCustRefLabel3Legend") & ":"
            If CBool(dr("StockOrderCustRef3Mandatory")) Then
                lblLegendCustRef3.ForeColor = Drawing.Color.Red
                rfvCustRef3.Enabled = True
                rfvCustRef3.EnableClientScript = True
            Else
                rfvCustRef3.Enabled = False
                rfvCustRef3.EnableClientScript = False
            End If
        Else
            trCustRef3.Visible = False
        End If
        If CBool(dr("StockOrderCustRef4Visible")) Then
            trCustRef4.Visible = True
            lblLegendCustRef4.Text = dr("StockOrderCustRefLabel4Legend") & ":"
            If CBool(dr("StockOrderCustRef4Mandatory")) Then
                lblLegendCustRef4.ForeColor = Drawing.Color.Red
                rfvCustRef4.Enabled = True
                rfvCustRef4.EnableClientScript = True
            Else
                rfvCustRef4.Enabled = False
                rfvCustRef4.EnableClientScript = False
            End If
        Else
            trCustRef3.Visible = False
        End If
        'lblAuthorisationAdvisory01.Text = dr("AuthorisationAdvisory")

        'trMultiAddressOrder.Visible = pbMultipleAddressOrders
    End Sub

    Protected Sub SetAddressVisibility(ByVal bVisible As Boolean)
        trPostCode.Visible = bVisible
        trCneeAddr1.Visible = bVisible
        trCneeAddr2.Visible = bVisible
        trCneeAddr3.Visible = bVisible
        trTownCity.Visible = bVisible
        trCneeState.Visible = bVisible
        trPostCode.Visible = bVisible
        'trCneeTel.Visible = bVisible
        'trCneeEmail.Visible = bVisible
    End Sub
                                      
    Protected Sub btnAddToOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If rcbProduct.SelectedValue = String.Empty Then
            Exit Sub
        End If
        
        'If Not IsNumeric(tbQty.Text) Then
        '    WebMsgBox.Show("Please specify a valid quantity.")
        '    tbQty.Focus()
        '    Exit Sub
        'Else
        '    If CInt(tbQty.Text) <= 0 Then
        '        WebMsgBox.Show("Please specify a positive non-zero quantity.")
        '        tbQty.Focus()
        '        Exit Sub
        '    End If
        'End If
        If Not IsNumeric(rntbQty.Text) Then
            WebMsgBox.Show("Please specify a valid quantity.")
            rntbQty.Focus()
            Exit Sub
        Else
            If CInt(rntbQty.Text) <= 0 Then
                WebMsgBox.Show("Please specify a positive non-zero quantity.")
                rntbQty.Focus()
                Exit Sub
            End If
        End If
        Call CreateBasketIfNull()
        gdtBasket = Session("BO_BasketData")
        Dim gdvBasket As New DataView(gdtBasket)
        gdvBasket.RowFilter = "LogisticProductKey='" & rcbProduct.SelectedValue & "'"
        If gdvBasket.Count > 0 Then
            'WebMsgBox.Show("This product (" & rcbProduct.SelectedItem.Text & ") is already in your basket.\n\nTo change the quantity, remove it from the basket and re-select the product with the quantity required.")
            WebMsgBox.Show("This product is already in your basket.\n\nTo change the quantity, remove it from the basket and re-select the product with the quantity required.")
        Else
            Dim nAvailableQty As Int32 = GetAvailableQty(rcbProduct.SelectedValue)
            'If nAvailableQty < CInt(tbQty.Text) Then
            If nAvailableQty < CInt(rntbQty.Text) Then
                If nAvailableQty = 0 Then
                    WebMsgBox.Show("None of these items is available.")
                Else
                    Dim sIsAre As String = "are"
                    If nAvailableQty = 1 Then
                        sIsAre = "is"
                    End If
                    WebMsgBox.Show("Only " & nAvailableQty.ToString & ") of these items " & sIsAre & " available. Sadly this is insufficient to fulfil your order.")
                End If
            Else
                Dim dr As DataRow
                dr = gdtBasket.NewRow()
                'dr("LogisticProductKey") = ddlProduct.SelectedValue
                dr("LogisticProductKey") = rcbProduct.SelectedValue
                'dr("Product") = ddlProduct.SelectedItem.Text
                'dr("Product") = rcbProduct.SelectedItem.Text
                'dr("Product") = rcbProduct.Text & Server.UrlEncode(" <i>(" & nAvailableQty.ToString & " available)</i>")
                dr("Product") = rcbProduct.Text
                dr("Available") = nAvailableQty.ToString
                'dr("Qty") = tbQty.Text
                dr("Qty") = rntbQty.Text
                gdtBasket.Rows.Add(dr)
            End If
        End If
        
        Session("BO_BasketData") = gdtBasket
        gvBasket.DataSource = gdtBasket
        gvBasket.DataBind()
        Call SetPlaceOrderButtonVisibility()
        'ddlProduct.SelectedIndex = 0
        rcbProduct.SelectedIndex = 0
        'tbQty.Text = String.Empty
        rntbQty.Text = "1"
        'tbQty.Text = "1"
        
        ddlCustomer.Enabled = False
        btnAddToOrder.Enabled = False
        'tbQty.Enabled = False
        rntbQty.Enabled = False
        'lnkbtnPlus1.Enabled = False
        'lnkbtnPlus5.Enabled = False
        'lnkbtnMinus1.Enabled = False
        'lnkbtnMinus5.Enabled = False
        'ddlProduct.Focus()
        rcbProduct.Focus()
        Call ClearOrderConfirmation()
        
        rcbProduct.Text = String.Empty
    End Sub

    Protected Sub SetPlaceOrderButtonVisibility()
        If gvBasket.Rows.Count > 0 Then
            btnPlaceOrder.Visible = True
        Else
            btnPlaceOrder.Visible = False
        End If
    End Sub
    
    Protected Sub CreateBasketIfNull()
        If IsNothing(Session("BO_BasketData")) Then
            gdtBasket = New DataTable()
            gdtBasket.Columns.Add(New DataColumn("Product", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("LogisticProductKey", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("Available", GetType(Long)))
            gdtBasket.Columns.Add(New DataColumn("Qty", GetType(Long)))
            Session("BO_BasketData") = gdtBasket
        End If
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
            'lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub GetCustomerAccountCodes()
        Dim olic As ListItemCollection
        olic = ExecuteQueryToListItemCollection("SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND ISNULL(AccountHandlerKey,0) > 0 ORDER BY CustomerAccountCode", "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Clear()
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In olic
            ddlCustomer.Items.Add(li)
        Next
    End Sub
   
    Protected Sub GetBookedByUsers()
        Dim olic As ListItemCollection
        olic = ExecuteQueryToListItemCollection("SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' 'Name', [key] 'UserKey' FROM UserProfile WHERE Status = 'ACTIVE' AND DeletedFlag = 0 AND CustomerKey = " & ddlCustomer.SelectedValue & " ORDER BY LastName", "Name", "UserKey")
        ddlBookedBy.Items.Clear()
        ddlBookedBy.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In olic
            ddlBookedBy.Items.Add(li)
        Next
    End Sub
   
    Protected Sub ddlCustomers_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCustomer.SelectedValue > 0 Then
            'Call PopulateProductDropdown(ddlCustomer.SelectedValue)			' XXXX
            pnImpersonateCustomer = ddlCustomer.SelectedValue
            Call GetBookedByUsers()
            ddlBookedBy.Enabled = True
            ddlBookedBy.Focus()
        Else
            ddlBookedBy.Enabled = False
        End If
        Call ClearOrderConfirmation()
    End Sub

    Protected Function GetAvailableQty(ByVal sLogisticProductKey As String) As Int32
        Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END"
        Try
            GetAvailableQty = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        Catch
            GetAvailableQty = 0
        End Try
    End Function
    
    'Protected Function GetProductsByCustomer(ByVal sCustomerKey As String, Optional ByVal sFilter As String = "") As DataTable            ' XXXX
    '    Dim sSQL As String
    '    sSQL = "SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription 'Product', LogisticProductKey, ThumbnailImage, Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation lpl ON lp.LogisticProductKey = lpl.LogisticProductKey WHERE ArchiveFlag = 'N' AND DeletedFlag = 'N' AND CustomerKey = " & sCustomerKey
    '    If sFilter <> String.Empty Then
    '        sFilter = sFilter.Replace("'", "''")
    '        sSQL += " AND (ProductCode LIKE '%" & sFilter & "%' OR ProductDescription LIKE '%" & sFilter & "%')"
    '    End If
    '    ' tbSearch.Text = tbSearch.Text.Trim
    '    If psSearchString <> String.Empty Then
    '        sSQL += " AND (ProductCode LIKE '%" & psSearchString & "%' OR ProductDescription LIKE '%" & psSearchString & "%')"
    '    End If
    '    sSQL += " ORDER BY ProductCode"
    '    Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
    '    GetProductsByCustomer = dt
    'End Function

    Protected Function GetProductsByCustomer(ByVal sCustomerKey As String, Optional ByVal sFilter As String = "") As DataTable            ' XXXX
        Dim dt As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_WUQuickOrder_GetProducts", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Filter", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Filter").Value = sFilter

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FavouriteProducts", SqlDbType.Bit))
        'oAdapter.SelectCommand.Parameters("@FavouriteProducts").Value = IIf(rbFavouriteProducts.Checked, 1, 0)
        oAdapter.SelectCommand.Parameters("@FavouriteProducts").Value = 0

        oAdapter.Fill(dt)
        
        GetProductsByCustomer = dt
    End Function
    

    'Protected Sub PopulateProductDropdown(ByVal sCustomerKey As String, Optional ByVal sFilter As String = "")
    '    Dim sSQL As String
    '    sSQL = "SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription 'Product', LogisticProductKey, ThumbnailImage FROM LogisticProduct WHERE ArchiveFlag = 'N' AND DeletedFlag = 'N' AND CustomerKey = " & sCustomerKey
    '    If sFilter <> String.Empty Then
    '        sSQL += " AND (ProductCode LIKE '%" & sFilter & "%' OR ProductDescription LIKE '%" & sFilter & "%')"
    '    End If
    '    sSQL += " ORDER BY ProductCode"
        
    '    rcbProduct.DataSource = Nothing
    '    rcbProduct.DataBind()
        
    '    Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
    '    rcbProduct.Items.Clear()
    '    If dt.Rows.Count = 0 Then
    '        If tbSearch.Text <> String.Empty Then
    '            rcbProduct.Items.Insert(0, New RadComboBoxItem("- no products match that filter -", 0))
    '        Else
    '            rcbProduct.Items.Insert(0, New RadComboBoxItem("- no products available -", 0))
    '        End If
    '        Exit Sub
    '    End If

    '    rcbProduct.DataSource = dt
    '    rcbProduct.DataTextField = "Product"
    '    rcbProduct.DataValueField = "LogisticProductKey"
    '    rcbProduct.DataBind()

    '    Dim sProductCount As String = dt.Rows.Count
    '    Dim sPlural As String = String.Empty
    '    If sProductCount <> 1 Then
    '        sPlural = "s"
    '    End If
    '    If tbSearch.Text <> String.Empty Then
    '        rcbProduct.Items.Insert(0, New RadComboBoxItem("- select from " & sProductCount & " product" & sPlural & " (search: " & tbSearch.Text & " ) -", 0))
    '    Else
    '        'rcbProduct.Items.Insert(0, New RadComboBoxItem("- select from " & sProductCount & " product" & sPlural & " (no search) -", 0))
    '        rcbProduct.Items.Insert(0, New RadComboBoxItem("- select from " & sProductCount & " product" & sPlural & " -", 0))
    '    End If
    'End Sub
    
    Protected Sub NormaliseCneeAddress()
        If ddlCountry.SelectedValue = COUNTRY_CODE_USA_NYC Then
            tbCneeState.Text = "NEW YORK CITY"
            Exit Sub
        End If
        If ddlUSStatesCanadianProvinces.SelectedIndex > 0 Then
            If ddlCountry.SelectedValue = COUNTRY_CODE_CANADA Or ddlCountry.SelectedValue = COUNTRY_CODE_USA Then
                tbCneeState.Text = ddlUSStatesCanadianProvinces.SelectedItem.Text
            End If
        End If
    End Sub
    
    Protected Function nSubmitConsignment() As Integer
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        Call NormaliseCneeAddress()
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        'lblError.Text = ""
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        'param1.Value = ddlBookedBy.SelectedValue
        param1.Value = pnImpersonateBookedByUser
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = pnImpersonateCustomer
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)
        
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)

        param3.Value = tbCustRef1.Text
        param4.Value = tbCustRef2.Text
        param5.Value = tbCustRef3.Text
        param6.Value = tbCustRef4.Text
        
        If IsArthritis() Then
            param3.Value = ddlArthritisCostCentre.SelectedValue
            param4.Value = ddlArthritisCategory.SelectedValue
            param5.Value = tbArthritisPONumber.Text
            param6.Value = String.Empty
        End If

        oCmdAddBooking.Parameters.Add(param3)
        oCmdAddBooking.Parameters.Add(param4)
        oCmdAddBooking.Parameters.Add(param5)
        oCmdAddBooking.Parameters.Add(param6)

        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = tbSpecialInstructions.Text.Replace(Environment.NewLine, " ").Trim
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = tbPackingNote.Text.Replace(Environment.NewLine, " ").Trim
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

        Dim dtCnor As DataTable = ExecuteQueryToDataTable("SELECT * FROM Customer WHERE CustomerKey = " & pnImpersonateCustomer)
        Dim drCnor As DataRow
        If dtCnor.Rows.Count = 1 Then
            drCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't find Consignor details.")
            Exit Function
        End If
       
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        'param13.Value = psCnorCompany
        param13.Value = drCnor("CustomerName")
       
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = drCnor("CustomerAddr1")
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = drCnor("CustomerAddr2")
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = drCnor("CustomerAddr3")
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = drCnor("CustomerTown")
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = drCnor("CustomerCounty")
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = drCnor("CustomerPostCode")
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = drCnor("CustomerCountryKey")
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = ""
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = ""
        oCmdAddBooking.Parameters.Add(param22)
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = ""
        oCmdAddBooking.Parameters.Add(param23)
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = tbCneeName.Text
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = tbCneeAddr1.Text
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = tbCneeAddr2.Text
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = tbCneeAddr3.Text
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = tbCneeTown.Text
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = tbCneeState.Text
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = tbCneePostCode.Text
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = ddlCountry.SelectedValue
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = tbCneeCtcName.Text
        oCmdAddBooking.Parameters.Add(param33)
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = tbCneeTel.Text
        oCmdAddBooking.Parameters.Add(param34)
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = tbCneeEmail.Text
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
        
        For i As Int32 = 0 To oCmdAddBooking.Parameters.Count - 1
            Trace.Write(oCmdAddBooking.Parameters(i).ParameterName.ToString)
            Trace.Write(oCmdAddBooking.Parameters(i).DbType.ToString)
            If Not IsNothing(oCmdAddBooking.Parameters(i).Value) Then
                Trace.Write(oCmdAddBooking.Parameters(i).Value.ToString)
            Else
                Trace.Write("NOTHING")
                
            End If
        Next
        
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
                gdtBasket = Session("BO_BasketData")
                For Each dr As DataRow In gdtBasket.Rows
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = pnImpersonateBookedByUser
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = pnImpersonateCustomer
                    oCmdAddStockItem.Parameters.Add(param52)
                    Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param53.Value = lBookingKey
                    oCmdAddStockItem.Parameters.Add(param53)
                    Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                    param54.Value = dr("LogisticProductKey")
                    oCmdAddStockItem.Parameters.Add(param54)
                    Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                    param55.Value = "PENDING"
                    oCmdAddStockItem.Parameters.Add(param55)
                    Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                    param56.Value = dr("Qty")
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
                'lblError.Text = "Error adding Web Booking [BookingKey=0]."
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                nSubmitConsignment = lConsignmentKey
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            'lblError.Text = ""
            'lblError.Text = ex.ToString
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function

    'Protected Sub ddlProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim ddl As DropDownList = sender
    '    If ddl.SelectedIndex = 0 Then
    '        btnAddToOrder.Enabled = False
    '        lnkbtnPlus1.Enabled = False
    '        lnkbtnPlus5.Enabled = False
    '        lnkbtnMinus1.Enabled = False
    '        lnkbtnMinus5.Enabled = False
    '        tbQty.Enabled = False
    '    Else
    '        btnAddToOrder.Enabled = True
    '        lnkbtnPlus1.Enabled = True
    '        lnkbtnPlus5.Enabled = True
    '        lnkbtnMinus1.Enabled = True
    '        lnkbtnMinus5.Enabled = True
    '        tbQty.Enabled = True
    '    End If
    '    tbQty.Focus()
    'End Sub

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

    Protected Sub btnPlaceOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'gdtBasket = Session("BO_BasketData")
        'If gdtBasket.Rows.Count = 0 Then
        '    WebMsgBox.Show("Your basket is empty.")
        '    Exit Sub
        'End If
       
        'Page.Validate()
        'If Not Page.IsValid Then
        '    WebMsgBox.Show("Please complete all required fields.")
        '    Exit Sub
        'End If
       
        'Dim sCheckQuantity As String = CheckSufficientQuantityAvailable()
        'If sCheckQuantity <> String.Empty Then
        '    WebMsgBox.Show(sCheckQuantity)
        '    Exit Sub
        'End If
        mpe.PopupControlID = "divConfirmOrder"
        mpe.Show()
        '    Call btnDummy_Click(Nothing, Nothing)
    End Sub
   
    Protected Function CheckSufficientQuantityAvailable() As String
        CheckSufficientQuantityAvailable = String.Empty
        gdtBasket = Session("BO_BasketData")
        For Each dr As DataRow In gdtBasket.Rows
            Dim nQtyAvailable As Int32 = GetQuantityAvailable(dr("LogisticProductKey"))
            If nQtyAvailable < dr("Qty") Then
                CheckSufficientQuantityAvailable = "Product """ & dr("Product") & """ has a requested quantity of " & dr("Qty") & " but only " & nQtyAvailable & " is/are available."
                Exit For
            End If
        Next
    End Function
   
    Protected Function GetQuantityAvailable(ByVal sLogisticProductKey As String) As Int32
        GetQuantityAvailable = 0
    End Function
   
    Protected Sub ClearUp()
        If divInternal.Visible Then
            ddlCustomer.SelectedIndex = 0
        End If
        tbCustRef1.Text = String.Empty
        tbCustRef2.Text = String.Empty
        tbCustRef3.Text = String.Empty
        tbCustRef4.Text = String.Empty
        tbCneeCtcName.Text = String.Empty
        tbCneeName.Text = String.Empty
        tbCneeAddr1.Text = String.Empty
        tbCneeAddr2.Text = String.Empty
        tbCneeAddr3.Text = String.Empty
        tbCneeTown.Text = String.Empty
        tbCneeState.Text = String.Empty
        tbCneePostCode.Text = String.Empty
        tbCneeTel.Text = String.Empty
        tbCneeEmail.Text = String.Empty
        ddlCountry.SelectedIndex = 0
        tbPackingNote.Text = String.Empty
        tbSpecialInstructions.Text = String.Empty
        rcbProduct.Text = String.Empty
        gdtBasket = Nothing
        Session("BO_BasketData") = gdtBasket
        If IsArthritis() Then
            ddlArthritisCategory.SelectedIndex = 0
            ddlArthritisCostCentre.SelectedIndex = 0
            tbArthritisPONumber.Text = String.Empty
        End If
    End Sub
   
    Protected Sub lnkbtnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' remove item
        Dim lb As LinkButton = sender
        Dim sLogisticProductKey As String = lb.CommandArgument
        Call RemoveItemFromBasket(sLogisticProductKey)
    End Sub
        
    Protected Sub RemoveItemFromBasket(sLogisticProductKey As String)
        gdtBasket = Session("BO_BasketData")
        Dim gdvBasketView = New DataView(gdtBasket)
        gdvBasketView.RowFilter = "LogisticProductKey='" & sLogisticProductKey & "'"
        If gdvBasketView.Count > 0 Then
            gdvBasketView.Delete(0)
        End If
        Session("BO_BasketData") = gdtBasket
        gvBasket.DataSource = gdtBasket
        gvBasket.DataBind()
        If gvBasket.Rows.Count > 0 Then
            btnPlaceOrder.Visible = True
        Else
            btnPlaceOrder.Visible = False
        End If
    End Sub
   
    Protected Sub btnPostcodeFind_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call FindAddress()
    End Sub

    Protected Sub FindAddress()
        tbCneePostCode.Text = tbCneePostCode.Text.Trim.ToUpper

        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult

        objInterimResults = objLookup.ByPostcode(tbCneePostCode.Text, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
       
        If objInterimResults.IsError OrElse objInterimResults.Results Is Nothing OrElse objInterimResults.Results.GetLength(0) = 0 Then
            lblLookupError.Visible = True
            lbLookupResults.Visible = False
            lblLookupError.Text = objInterimResults.ErrorMessage
            If lblLookupError.Text.Trim = String.Empty Then
                lblLookupError.Text = "<br />No results found for this post code"
            Else
                lblLookupError.Text = "<br />" & lblLookupError.Text
            End If
            trPostcodeLookupOutput.Visible = False
            tbCneePostCode.Focus()
        Else
            lblLookupError.Visible = False
            lbLookupResults.Visible = True

            lbLookupResults.Items.Clear()

            If Not objInterimResults.Results Is Nothing Then
                For Each objInterimResult In objInterimResults.Results
                    lbLookupResults.Items.Add(New ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            trPostcodeLookupOutput.Visible = True
            lbLookupResults.Focus()
        End If
    End Sub
   
    Protected Sub lbLookupResults_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objAddressResults As uk.co.postcodeanywhere.services.AddressResults
        Dim objAddress As uk.co.postcodeanywhere.services.Address

        objAddressResults = objLookup.FetchAddress(lbLookupResults.SelectedValue, _
           uk.co.postcodeanywhere.services.enLanguage.enLanguageEnglish, _
           uk.co.postcodeanywhere.services.enContentType.enContentStandardAddress, _
           ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        If objAddressResults.IsError Then
            lblLookupError.Text = objAddressResults.ErrorMessage
        Else
            objAddress = objAddressResults.Results(0)

            'txtCneeCtcName.Text = objAddress.OrganisationName
            tbCneeName.Text = objAddress.OrganisationName.Trim
            tbCneeAddr1.Text = objAddress.Line1
            tbCneeAddr2.Text = objAddress.Line2
            tbCneeAddr3.Text = objAddress.Line3
            tbCneeTown.Text = objAddress.PostTown
            tbCneePostCode.Text = objAddress.Postcode
            tbCneeState.Text = objAddress.County

        End If
        trPostcodeLookupOutput.Visible = False
        If tbCneeName.Text = String.Empty Then
            tbCneeName.Focus()
        Else
            tbCneeCtcName.Focus()
        End If
    End Sub

    Protected Sub lnkbtnCancelPostcodeLookup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lbLookupResults.Items.Clear()
        lblLookupError.Visible = False
        trPostcodeLookupOutput.Visible = False
        tbCneePostCode.Focus()
    End Sub
   
    Protected Sub lnkbtnUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 1 To ddlCountry.Items.Count - 1
            If ddlCountry.Items(i).Text = "UK" Or ddlCountry.Items(i).Text = "U.K." Then
                ddlCountry.SelectedIndex = i
                Call SetAddressVisibility(True)
                btnPostcodeFind.Visible = True
                tbCneePostCode.Focus()
                Exit For
            End If
        Next
        Call HideCountryRelatedControls()
        Call SetCountryOther()
    End Sub
   
    Protected Sub ddlCountry_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCountry.SelectedValue > 0 Then
            If ddlCountry.SelectedValue = COUNTRY_UK Then
                btnPostcodeFind.Visible = True
            Else
                btnPostcodeFind.Visible = False
            End If
            Call SetAddressVisibility(True)
            tbCneePostCode.Focus()
        Else
            Call SetAddressVisibility(False)
        End If
        Call SetCountry(ddlCountry.SelectedValue, "")
        tbCneePostCode.Text = String.Empty
    End Sub

    Protected Sub SetCountryFieldsVisibility()
        Call HideCountryRelatedControls()
        If ddlCountry.SelectedValue = COUNTRY_CODE_USA Then
            trUSStatesCanadianProvinces.Visible = True
            trUSStateShortcuts.Visible = True
        ElseIf ddlCountry.SelectedValue = COUNTRY_CODE_USA_NYC Then
            trUSStatesCanadianProvinces.Visible = True
            trUSStateShortcuts.Visible = True
        ElseIf ddlCountry.SelectedValue = COUNTRY_CODE_CANADA Then
            trUSStatesCanadianProvinces.Visible = True
        Else
            trCneeState.Visible = True
        End If
    End Sub
    
    Protected Sub SetCountry(ByVal nCountryKey As Int32, ByVal sStateOrProvince As String)
        If nCountryKey = COUNTRY_CODE_USA Then
            Call SetCountryUSA(sStateOrProvince)
            'trUSStatesCanadianProvinces.Visible = True
            'trUSStateShortcuts.Visible = True
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
        'tbCneeState.Visible = True
        trCneeState.Visible = True
        lblLegendCountyStateRegionProvince.Text = "County / Region:"

        Dim l As Label
        lblLegendCountyStateRegionProvince.ForeColor = Drawing.Color.Black
        tbCneeState.Text = String.Empty
        lblLegendCountyStateRegionProvince.Font.Bold = False
        'rfvRegion.Enabled = False
        lblLegendPostcode.Text = "Post Code:"
    End Sub
   
    Protected Sub SetCountryUSA(ByVal sState As String)
        Call HideCountryRelatedControls()
        trUSStatesCanadianProvinces.Visible = True
        trUSStateShortcuts.Visible = True
        'ddlUSStatesCanadianProvinces.Visible = True
        lblLegendStateProvince.Text = "State:"
        'lblLegendCountyStateRegionProvince.Text = "State"
        lblLegendCountyStateRegionProvince.ForeColor = Drawing.Color.Red
        Call PopulateUSStatesDropdown()
        If sState <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text = sState Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        'rfvRegion.Enabled = True
        tbCneeState.Text = String.Empty
        lblLegendCountyStateRegionProvince.Font.Bold = True
        lblLegendPostcode.Text = "Zip Code:"
    End Sub
   
    Protected Sub SetCountryUSANewYorkCity()
        Call HideCountryRelatedControls()
        trUSStatesCanadianProvinces.Visible = True
        trUSStateShortcuts.Visible = True
        'lblLegendNewYorkCity.Visible = True
        'lblLegendCountyStateRegionProvince.Text = "State"
        lblLegendStateProvince.Text = "State:"
        'lblLegendCountyStateRegionProvince.ForeColor = Drawing.Color.Red
        'rfvRegion.Enabled = False
        '        tbCneeState.Text = lblLegendNewYorkCity.Text
        lblLegendPostcode.Text = "Zip Code:"
    End Sub
   
    Protected Sub SetCountryCanada(ByVal sProvince As String)
        Call HideCountryRelatedControls()
        trUSStatesCanadianProvinces.Visible = True
        'ddlUSStatesCanadianProvinces.Visible = True
        'lblLegendCountyStateRegionProvince.Text = "Province"
        lblLegendStateProvince.Text = "Province:"
        'lblLegendCountyStateRegionProvince.ForeColor = Drawing.Color.Red
        Call PopulateCanadianProvincesDropdown()
        If sProvince <> String.Empty Then
            For i As Int32 = 0 To ddlUSStatesCanadianProvinces.Items.Count - 1
                If ddlUSStatesCanadianProvinces.Items(i).Text = sProvince Then
                    ddlUSStatesCanadianProvinces.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        'rfvRegion.Enabled = True
        tbCneeState.Text = String.Empty
        lblLegendCountyStateRegionProvince.Font.Bold = True
        lblLegendPostcode.Text = "Postal Code:"
    End Sub
   
    Protected Sub HideCountryRelatedControls()
        trUSStatesCanadianProvinces.Visible = False
        trUSStateShortcuts.Visible = False
        trCneeState.Visible = False
        'ddlUSStatesCanadianProvinces.Visible = False
        'lblLegendNewYorkCity.Visible = False
        'tbCneeState.Visible = False
        trUSStateShortcuts.Visible = False
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

    'Protected Sub ddlUSStatesCanadianProvinces_SelectedIndexChanged(sender As Object, e As System.EventArgs)
    '    If ddlUSStatesCanadianProvinces.SelectedIndex > 0 Then
    '        tbCneeState.Text = ddlUSStatesCanadianProvinces.SelectedItem.Text
    '    Else
    '        tbCneeState.Text = String.Empty
    '    End If
    'End Sub

    Protected Sub lnkbtnNewYorkCity_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 1 To ddlUSStatesCanadianProvinces.Items.Count - 1
            If ddlUSStatesCanadianProvinces.Items(i).Text.ToLower.Contains("new york") Then
                ddlUSStatesCanadianProvinces.SelectedIndex = i
                Exit For
            End If
        Next
        tbCneeTown.Text = "New York City"
    End Sub

    Protected Sub lnkbtnWashingtonDC_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbCneeTown.Text = "Washington D.C."
    End Sub
    
    Property pnImpersonateCustomer() As Int32
        Get
            Dim o As Object = ViewState("QO_ImpersonateCustomer")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("QO_ImpersonateCustomer") = Value
        End Set
    End Property

    Property pnImpersonateBookedByUser() As Int32
        Get
            Dim o As Object = ViewState("QO_ImpersonateBookedByUser")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("QO_ImpersonateBookedByUser") = Value
        End Set
    End Property
    
    Property psSearchString() As String
        Get
            Dim o As Object = ViewState("QO_SearchString")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("QO_SearchString") = Value
        End Set
    End Property
  
    Protected Sub ddlBookedBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue > 0 Then
            pnImpersonateBookedByUser = ddl.SelectedValue
            divMainForm.Visible = True
            ddlCountry.Focus()
        Else
            divMainForm.Visible = False
        End If
        Call ClearOrderConfirmation()
    End Sub
    
    Protected Sub lnkbtnShowHideAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
    
    Protected Sub lnkbtnAlterQuantity_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim lnkbtn As LinkButton = sender
        'Dim nQty As Int32 = CInt(lnkbtn.CommandArgument)
        'tbQty.Text = tbQty.Text.Trim
        'tbQty.Text = tbQty.Text.TrimStart("0")
        'If tbQty.Text = String.Empty Then
        '    tbQty.Text = "0"
        'End If
        'If tbQty.Text.Length < 8 AndAlso IsNumeric(tbQty.Text) Then
        '    tbQty.Text = tbQty.Text + nQty
        '    If CInt(tbQty.Text) < 0 Then
        '        tbQty.Text = "1"
        '    End If
        'End If
    End Sub
    
    Protected Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nConsignmentNumber As Int32 = nSubmitConsignment()
        If nConsignmentNumber > 0 Then
            trConsignmentNumber.Visible = True
            lblConsignmentNumber.Text = nConsignmentNumber.ToString
            gvBasket.DataSource = Nothing
            gvBasket.DataBind()
            btnPlaceOrder.Visible = False
            'lblConsignment.Text = "Consignment # " & nConsignmentNumber
        Else
            mpe.PopupControlID = "divOrderError"
            mpe.Show()

            'lblConsignment.Text = "COULD NOT CREATE CONSIGNMENT"
        End If
        Call ClearUp()
    End Sub
    
    Protected Sub ClearOrderConfirmation()
        lblConsignmentNumber.Text = String.Empty
        trConsignmentNumber.Visible = False
        'If tbSearch.Text <> String.Empty Then
        '    tbSearch.Text = String.Empty
        '    'Call PopulateProductDropdown(pnImpersonateCustomer)		 ' XXXX
        'End If
    End Sub
    
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim x As Int32 = 1
    End Sub
    
    Property psVirtualThumbURL() As String
        Get
            Dim o As Object = ViewState("QO_VirtualThumbURL")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("QO_VirtualThumbURL") = Value
        End Set
    End Property
    
    Protected Function GetFullImageURL(ByVal DataItem As Object) As String
        GetFullImageURL = "http://my.transworld.eu.com/common/prod_images/thumbs/" & DataBinder.Eval(DataItem, "ThumbnailImage")
    End Function
  
    Protected Sub rcbProduct_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)          ' XXXX
        Dim s As String = e.Text
        Dim data As DataTable = GetProductsByCustomer(pnImpersonateCustomer, e.Text)
        Dim sThumbnailImage As String
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbProduct.DataTextField = "Product"
        rcbProduct.DataValueField = "LogisticProductKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            rcb.Text = data.Rows(i)("Product").ToString()
            rcb.Value = data.Rows(i)("LogisticProductKey").ToString()
            sThumbnailImage = data.Rows(i)("ThumbnailImage").ToString()
            rcbProduct.Items.Add(rcb)
            Dim lblProduct As Label = rcb.FindControl("lblProduct")
            Dim imgProduct As Image = rcb.FindControl("imgProduct")
            lblProduct.Text = data.Rows(i)("Product").ToString()
            imgProduct.ImageUrl = "http://my.transworld.eu.com/common/prod_images/thumbs/" & data.Rows(i)("ThumbnailImage").ToString()
        Next
        e.Message = GetStatusMessage(endOffset, data.Rows.Count)
    End Sub

    Private Shared Function GetStatusMessage(ByVal nOffset As Integer, ByVal nTotal As Integer) As String
        If nTotal <= 0 Then
            Return "No matches"
        End If
        'Return [String].Format("Items <b>1</b>-<b>{0}</b> of <b>{1}</b>", nOffset, nTotal)
        If nOffset <= ITEMS_PER_REQUEST Then
            'GetStatusMessage = "Click for more items " & nOffset & " " & nTotal & " " & ITEMS_PER_REQUEST
            GetStatusMessage = "Click for more items"
        End If
        'GetStatusMessage = "Click for more items" 
        If nOffset = nTotal Then
            GetStatusMessage = "No more items"
        Else
            GetStatusMessage = "Click for more items"
        End If
    End Function
    
    Protected Sub rcbProduct_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        'Dim ddl As DropDownList = sender
        If rcb.SelectedIndex = 0 Then
            btnAddToOrder.Enabled = False
            'lnkbtnPlus1.Enabled = False
            'lnkbtnPlus5.Enabled = False
            'lnkbtnMinus1.Enabled = False
            'lnkbtnMinus5.Enabled = False
            rntbQty.Enabled = False
            'tbQty.Enabled = False
        Else
            btnAddToOrder.Enabled = True
            'lnkbtnPlus1.Enabled = True
            'lnkbtnPlus5.Enabled = True
            'lnkbtnMinus1.Enabled = True
            'lnkbtnMinus5.Enabled = True
            'tbQty.Enabled = True
            'tbQty.Focus()
            rntbQty.Enabled = True
            rntbQty.Focus()
        End If
        Call ClearOrderConfirmation()
    End Sub
    
    Protected Sub lnkbtnPlaceAnotherOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearOrderConfirmation()
        rcbProduct.Focus()
    End Sub
    
    'Protected Sub btnSearchGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    tbSearch.Text = tbSearch.Text.Trim
    '    If tbSearch.Text <> String.Empty Then
    '        'Call PopulateProductDropdown(pnImpersonateCustomer, tbSearch.Text)		 ' XXXX
    '        rcbProduct.Focus()
    '    End If
    '    psSearchString = tbSearch.Text
    'End Sub
    
    Protected Sub lnkbtnClearSearchTerm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearSearchTerm()
    End Sub
    
    Protected Sub ClearSearchTerm()
        ' tbSearch.Text = String.Empty
        'Call PopulateProductDropdown(pnImpersonateCustomer)					   ' XXXX
        rcbProduct.Focus()
        psSearchString = String.Empty
    End Sub
    
    'Protected Sub cbFilterProducts_CheckedChanged(sender As Object, e As System.EventArgs)
    '    Dim cb As CheckBox = sender
    '    Call SetFilterControlsVisibility(cb.Checked)
    'End Sub
    
    'Protected Sub SetFilterControlsVisibility(bVisible As Boolean)
    '    tbSearch.Visible = bVisible
    '    btnSearchGo.Visible = bVisible
    '    lnkbtnClearSearchTerm.Visible = bVisible
    '    If bVisible Then
    '        tbSearch.Focus()
    '    Else
    '        rcbProduct.Focus()
    '        Call ClearSearchTerm()
    '    End If
    'End Sub
    
    Protected Sub lnkbtnRemove_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Dim imgbtn As ImageButton = sender
        Dim sLogisticProductKey As String = imgbtn.CommandArgument
        Call RemoveItemFromBasket(sLogisticProductKey)
    End Sub
    
    Protected Sub SetAddressFieldsVisibility(bVisible As Boolean)
        trCountry.Visible = bVisible
        trPostCode.Visible = bVisible
        If Not bVisible Then
            trPostcodeLookupOutput.Visible = bVisible
        End If
        trCneeAddr1.Visible = bVisible
        trCneeAddr2.Visible = bVisible
        trCneeAddr3.Visible = bVisible
        trTownCity.Visible = bVisible
        trCneeState.Visible = bVisible
        trUSStatesCanadianProvinces.Visible = bVisible
        trUSStateShortcuts.Visible = bVisible
        'trCneeTel.Visible = bVisible
        'trCneeEmail.Visible = bVisible
    End Sub

    Protected Sub BindAddressBook()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDT As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddresses", oConn)
        Dim sSearchCriteria As String = tbSearchAddressBook.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        'lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UseSharedAddressBook", SqlDbType.Bit))
            oAdapter.SelectCommand.Parameters("@UseSharedAddressBook").Value = rbSharedAddressBook.Checked
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FieldMask", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@FieldMask").Value = 0                      ' ddlAddressFields.SelectedValue  ' 0=all fields, 1=Company Name
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@UserKey").Value = pnImpersonateBookedByUser

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = pnImpersonateCustomer

            oAdapter.Fill(oDT)
            gvAddressBook.DataSource = oDT
            gvAddressBook.DataBind()
        Catch ex As SqlException
            'lblError.Text = ""
            'lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
        
    Protected Sub SetAddressBookFieldsVisibility(bVisible As Boolean)
        trAddressBook01.Visible = bVisible
        trAddressBook02.Visible = bVisible
        trAddressBook03.Visible = bVisible
        trAddressBook04.Visible = bVisible
    End Sub

    Protected Sub lnkbtnShowAddressBook_Click(sender As Object, e As System.EventArgs)
        If lnkbtnShowAddressBook.Text.ToLower.Contains("show") Then
            Call SetAddressFieldsVisibility(False)
            Call SetAddressBookFieldsVisibility(True)
            Call BindAddressBook()
            btnPlaceOrder.Visible = False
            lnkbtnShowAddressBook.Text = "hide address book"
            Call SetProductSelectability(False)
        Else
            lnkbtnShowAddressBook.Text = "show address book"
            Call SetAddressBookFieldsVisibility(False)
            Call SetAddressFieldsVisibility(True)
            Call SetCountryFieldsVisibility()
            Call SetPlaceOrderButtonVisibility()
            Call SetProductSelectability(True)
        End If
    End Sub
    
    Protected Sub SetProductSelectability(bEnabled As Boolean)
        rcbProduct.Enabled = bEnabled
        'cbFilterProducts.Enabled = bEnabled
        'btnSearchGo.Enabled = bEnabled
        'lnkbtnClearSearchTerm.Enabled = bEnabled
    End Sub
    
    Protected Sub btnSelectAddress_Click(sender As Object, e As System.EventArgs)
        Dim btn As Button = sender
        Dim nAddressKey As Integer = btn.CommandArgument
        Call SetAddressBookFieldsVisibility(False)
        Call SetAddressFieldsVisibility(True)
        Call GetConsigneeAddress(nAddressKey)
        Call SetPlaceOrderButtonVisibility()
        Call SetProductSelectability(True)
        lnkbtnShowAddressBook.Text = "show address book"
    End Sub
    
    Protected Sub btnSearchAddressBook_Click(sender As Object, e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub gvAddressBook_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvAddressBook.PageIndex = e.NewPageIndex
        Call BindAddressBook()
    End Sub
    
    Protected Sub rbPersonalAddressBook_CheckedChanged(sender As Object, e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub

    Protected Sub rbSharedAddressBook_CheckedChanged(sender As Object, e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub GetConsigneeAddress(nAddressKey As Int32)
        If nAddressKey > 0 Then
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_GlobalAddress_GetFromKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@DestKey", SqlDbType.Int, 4)
            oCmd.Parameters.Add(oParam)
            oParam.Value = nAddressKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                tbCneeName.Text = oDataReader("Company").ToString.Trim & String.Empty
                tbCneeAddr1.Text = oDataReader("Addr1").ToString.Trim & String.Empty
                tbCneeAddr2.Text = oDataReader("Addr2").ToString.Trim & String.Empty
                tbCneeAddr3.Text = oDataReader("Addr3").ToString.Trim & String.Empty
                tbCneeTown.Text = oDataReader("Town").ToString.Trim & String.Empty
                tbCneeState.Text = oDataReader("State").ToString.Trim & String.Empty
                tbCneePostCode.Text = oDataReader("PostCode").ToString.Trim & String.Empty
                tbCneeTel.Text = oDataReader("Telephone").ToString.Trim & String.Empty
                tbCneeEmail.Text = oDataReader("Email").ToString.Trim & String.Empty

                If Not IsDBNull(oDataReader("CountryKey")) Then
                    ' NEXT TWO CALLS MAY BE CALLING SOME METHODS TWICE
                    Call SetCountryDropdown(oDataReader("CountryKey"))
                    Call SetCountry(oDataReader("CountryKey"), oDataReader("State").ToString.Trim & String.Empty)
                End If
                tbCneeCtcName.Text = oDataReader("AttnOf").ToString.Trim
                oDataReader.Close()
            Catch ex As SqlException
                '    lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            'Call SaveRetrievedAddress()
            '
            ' DO SOMETHING SENSIBLE HERE FOR US / CANADA
            '
            ' ...but for now...
            Call SetCountryFieldsVisibility()
            'trUSStatesCanadianProvinces.Visible = False
            'trPostcodeLookupOutput.Visible = False
            'trUSStateShortcuts.Visible = False
            If ddlCountry.SelectedValue = 222 Then
                btnPostcodeFind.Visible = True
            Else
                btnPostcodeFind.Visible = False
            End If
        End If
    End Sub
    
    Protected Sub SetCountryDropdown(ByVal sCountryKey As String)
        If IsNumeric(sCountryKey) Then
            Dim nCountryKey As Integer = CInt(sCountryKey)
            For i As Integer = 0 To ddlCountry.Items.Count - 1
                If ddlCountry.Items(i).Value = nCountryKey Then
                    ddlCountry.SelectedIndex = i
                    Call SetCountry(ddlCountry.SelectedValue, "")
                    Exit For
                End If
            Next
        End If
    End Sub
    

    Protected Sub lnkbtnClearFilter_Click(sender As Object, e As System.EventArgs)
        rcbProduct.Text = String.Empty
        btnAddToOrder.Enabled = False
        rcbProduct.Focus()
    End Sub
    
    Protected Sub lnkbtnClearAddressBookSearch_Click(sender As Object, e As System.EventArgs)
        tbSearchAddressBook.Text = String.Empty
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Quick Order</title>
    <style type="text/css">
        .style1
        {
            height: 35px;
        }
        .style2
        {
            height: 23px;
        }
        .style3
        {
            height: 28px;
        }
        .qtychange
        {
            text-decoration: none;
            color: #999999;
        }
        .qtychange:LINK
        {
            text-decoration: none;
            color: #999999;
        }
        .qtychange:ACTIVE
        {
            text-decoration: none;
            color: #999999;
        }
        .qtychange:VISITED
        {
            text-decoration: none;
            color: #999999;
        }
        .qtychange:HOVER
        {
            text-decoration: none;
            color: #999999;
        }
        A.qtychange:LINK
        {
            text-decoration: none;
            color: #999999;
        }
        A.qtychange:VISITED
        {
            text-decoration: none;
            color: #999999;
        }
        A.qtychange:HOVER
        {
            text-decoration: none;
            color: #999999;
        }
    </style>
    <link href="css/modalpopup.css" rel="stylesheet" type="text/css" />
</head>
<body style="font-size: xx-small; font-family: Verdana">
    <form id="frmOrder" runat="server">
    <asp:ScriptManager runat="server" />
    <main:Header ID="ctlHeader" runat="server" />
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_addressbook">
            <td style="width: 50%; white-space: nowrap">
                &nbsp;
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                &nbsp;
            </td>
        </tr>
    </table>
    <div id="divInternal" runat="server" visible="false">
        <br />
        <table>
            <tr>
                <td style="width: 110px" />
                <td style="width: 5px" />
                <td style="width: 300px" />
            </tr>
            <tr>
                <td align="right">
                    <asp:CompareValidator ID="cvCustomer" runat="server" ControlToValidate="ddlCustomer"
                        Font-Names="Verdana" Operator="NotEqual" ValueToCompare="0" Text="###" Font-Bold="True" />
                    &nbsp;<asp:Label ID="Label1" runat="server" Text="Customer:"></asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:DropDownList ID="ddlCustomer" runat="server" Width="100%" OnSelectedIndexChanged="ddlCustomers_SelectedIndexChanged"
                        AutoPostBack="True" Font-Size="XX-Small">
                    </asp:DropDownList>
                    <br />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:CompareValidator ID="cvCneeCountry0" runat="server" ControlToValidate="ddlBookedBy"
                        Font-Names="Verdana" Operator="NotEqual" ValueToCompare="0" Text="###" Font-Bold="True" />
                    &nbsp;<asp:Label ID="Label13" runat="server" Text="Booked By:" />
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:DropDownList ID="ddlBookedBy" runat="server" Width="100%" Font-Size="XX-Small"
                        OnSelectedIndexChanged="ddlBookedBy_SelectedIndexChanged" AutoPostBack="True"
                        Enabled="False" />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
        </table>
    </div>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr valign="top">
            <td style="width: 50%">
                <table>
                    <tr>
                        <td style="width: 65px" />
                        <td style="width: 550px">
                            <asp:Label ID="Label12" runat="server" Text="Products" Font-Bold="True" Font-Names="Arial"
                                Font-Size="Small" />
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:LinkButton ID="lnkbtnClearFilter" runat="server" OnClick="lnkbtnClearFilter_Click"
                                CausesValidation="False" ToolTip="Clears product search text, if present." 
                                Font-Names="Arial" Font-Size="XX-Small">clear search</asp:LinkButton>
                        </td>
                    </tr>
                    <%--                    <tr runat="server" visible="false">
                        <td>
                        </td>
                        <td>
                            <asp:Panel ID="pnlSearch" runat="server" DefaultButton="btnSearchGo">
                                <asp:CheckBox ID="cbFilterProducts" runat="server" AutoPostBack="True" OnCheckedChanged="cbFilterProducts_CheckedChanged"
                                    Text="Search Products" />
                                &nbsp;&nbsp;
                                <asp:TextBox ID="tbSearch" runat="server" Font-Size="XX-Small" Width="150px" />
                                &nbsp;<asp:Button ID="btnSearchGo" runat="server" CausesValidation="false" OnClick="btnSearchGo_Click"
                                    Text="go" />
                                &nbsp;<asp:LinkButton ID="lnkbtnClearSearchTerm" runat="server" Font-Names="Arial"
                                    Font-Size="XX-Small" OnClick="lnkbtnClearSearchTerm_Click" CausesValidation="false">clear search</asp:LinkButton>
                            </asp:Panel>
                        </td>
                    </tr>
--%>                    <tr>
                        <td>
                        </td>
                        <td>
                            <telerik:RadComboBox ID="rcbProduct" runat="server" Width="300px" Font-Names="Arial"
                                Font-Size="X-Small" Font-Bold="true" OnSelectedIndexChanged="rcbProduct_SelectedIndexChanged"
                                AutoPostBack="True" HighlightTemplatedItems="true" CausesValidation="False" EnableLoadOnDemand="True"
                                OnItemsRequested="rcbProduct_ItemsRequested" EnableVirtualScrolling="True" ShowMoreResultsBox="True"
                                Filter="Contains" ToolTip="Shows all available products when no search text is specified. Search for products by typing a product code or description.">
                                <ItemTemplate>
                                    <table>
                                        <tr>
                                            <td style="width: 70px">
                                                <asp:Image ID="imgProduct" runat="server" />
                                            </td>
                                            <td style="width: 220px">
                                                <asp:Label ID="lblProduct" runat="server" />
                                            </td>
                                        </tr>
                                    </table>
                                </ItemTemplate>
                            </telerik:RadComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                        <td>
                            <asp:Label ID="Label560" runat="server" Text="Qty:" />
                            &nbsp;&nbsp;<telerik:RadNumericTextBox ID="rntbQty" runat="server" Font-Bold="True"
                                Font-Names="Arial" Font-Size="X-Small" MaxValue="100000" MinValue="1" ShowSpinButtons="True"
                                Value="1" Width="50px">
                                <NumberFormat DecimalDigits="0" />
                            </telerik:RadNumericTextBox>
                            &nbsp;<asp:Button ID="btnAddToOrder" runat="server" Text="Add to Order" Width="169px"
                                OnClick="btnAddToOrder_Click" CausesValidation="False" Enabled="False" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                        <td>
                        </td>
                    </tr>
                </table>
                <table>
                    <tr>
                        <td style="width: 65px" />
                        <td style="width: 550px" />
                        <asp:Label ID="Label557" runat="server" Text="Basket" Font-Bold="True" Font-Names="Arial"
                            Font-Size="Small" />
                    </tr>
                    <tr>
                        <td />
                        <td>
                            <asp:GridView ID="gvBasket" runat="server" CellPadding="2" Width="100%" AutoGenerateColumns="False">
                                <AlternatingRowStyle BackColor="#FFFF99" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgBtnRemove" runat="server" ImageUrl="~/images/delete.gif"
                                                CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' OnClick="lnkbtnRemove_Click"
                                                CausesValidation="false" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Product" HeaderText="Product" ReadOnly="True" />
                                    <asp:BoundField DataField="Available" HeaderText="Available">
                                        <ItemStyle ForeColor="#999999" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Qty" HeaderText="Order Qty" ReadOnly="True" SortExpression="Qty" />
                                </Columns>
                                <EmptyDataRowStyle BackColor="#FFFFCC" />
                                <EmptyDataTemplate>
                                    your basket is empty
                                </EmptyDataTemplate>
                                <RowStyle BackColor="#FFFFCC" />
                            </asp:GridView>
                        </td>
                    </tr>
                    <tr>
                        <td />
                        <td>
                            <asp:Button ID="btnPlaceOrder" runat="server" Text="Place Order" Width="150px" OnClick="btnPlaceOrder_Click"
                                Visible="False" />
                            &nbsp;
                        </td>
                    </tr>
                    <tr id="trConsignmentNumber" runat="server" visible="false">
                        <td class="style3" />
                        <td class="style3">
                            <asp:Label ID="lblOrderPlaced" runat="server" Text="Thank you for your order. Please note your consignment number:"
                                Font-Size="X-Small" />
                            <br />
                            <br />
                            &nbsp;
                            <asp:Label ID="lblConsignmentNumber" runat="server" Font-Bold="True" Font-Names="Arial"
                                Font-Size="Small" ForeColor="Blue" />
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:LinkButton ID="lnkbtnPlaceAnotherOrder" runat="server" OnClick="lnkbtnPlaceAnotherOrder_Click"
                                CausesValidation="False">place another order</asp:LinkButton>
                        </td>
                    </tr>
                </table>
            </td>
            <td style="width: 50%">
                <table>
                    <tr>
                        <td style="width: 110px" />
                        <td style="width: 5px" />
                        <td style="width: 300px" />
                        <asp:Label ID="Label555" runat="server" Text="Destination" Font-Bold="True" Font-Names="Arial"
                            Font-Size="Small" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnShowAddressBook" runat="server" OnClick="lnkbtnShowAddressBook_Click"
                            CausesValidation="False" Font-Names="Arial" Font-Size="XX-Small">show address book</asp:LinkButton>
                    </tr>
                    <tr id="trAddressBook01" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label561" runat="server" Text="Address Book:" />
                        </td>
                        <td />
                        <td>
                            <asp:RadioButton ID="rbPersonalAddressBook" runat="server" GroupName="addressbook"
                                Text="Personal" AutoPostBack="True" OnCheckedChanged="rbPersonalAddressBook_CheckedChanged" />
                            <asp:RadioButton ID="rbSharedAddressBook" runat="server" GroupName="addressbook"
                                Text="Shared" Checked="True" 
                                OnCheckedChanged="rbSharedAddressBook_CheckedChanged" AutoPostBack="True" />
                        </td>
                    </tr>
                    <tr id="trAddressBook02" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label562" runat="server" Text="Search:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbSearchAddressBook" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                            &nbsp;<asp:Button ID="btnSearchAddressBook" runat="server" Text="go" CausesValidation="False"
                                OnClick="btnSearchAddressBook_Click" />
                        &nbsp;&nbsp;&nbsp;
                            <asp:LinkButton ID="lnkbtnClearAddressBookSearch" runat="server" 
                                CausesValidation="False" Font-Names="Arial" Font-Size="XX-Small" 
                                onclick="lnkbtnClearAddressBookSearch_Click">clear search</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trAddressBook03" runat="server" visible="false">
                        <td align="right" colspan="3">
                            <asp:GridView ID="gvAddressBook" runat="server" Width="100%" CellPadding="2" Font-Names="Arial"
                                Font-Size="XX-Small" AllowPaging="True" AutoGenerateColumns="False" OnPageIndexChanging="gvAddressBook_PageIndexChanging">
                                <AlternatingRowStyle BackColor="#CCFFFF" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:Button ID="btnSelectAddress" CommandArgument='<%# Container.DataItem("DestKey")%>'
                                                runat="server" Text="select" OnClick="btnSelectAddress_Click" CausesValidation="False" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="AttnOf" HeaderText="Attn" ReadOnly="True" SortExpression="AttnOf" />
                                    <asp:BoundField DataField="Company" HeaderText="Name" ReadOnly="True" SortExpression="Company" />
                                    <asp:BoundField DataField="Addr1" HeaderText="Addr 1" ReadOnly="True" SortExpression="Addr1" />
                                    <asp:BoundField DataField="Town" HeaderText="Town" ReadOnly="True" SortExpression="Town" />
                                    <asp:BoundField DataField="CountryName" HeaderText="Country" ReadOnly="True" SortExpression="CountryName" />
                                </Columns>
                                <EmptyDataTemplate>
                                    no addresses found
                                </EmptyDataTemplate>
                                <PagerSettings Mode="NumericFirstLast" />
                                <RowStyle BackColor="#CCFFCC" />
                            </asp:GridView>
                        </td>
                    </tr>
                    <tr id="trAddressBook04" runat="server" visible="false">
                        <td align="right">
                        </td>
                        <td />
                        <td>
                        </td>
                    </tr>
                    <tr id="trCountry" runat="server">
                        <td align="right">
                            <asp:CompareValidator ID="cvCneeCountry" runat="server" ControlToValidate="ddlCountry"
                                Font-Names="Verdana" Operator="NotEqual" ValueToCompare="0" Text="###" Font-Size="XX-Small"
                                Font-Bold="True" />
                            &nbsp;<asp:Label ID="Label6" runat="server" Text="Country:" ForeColor="Red" />
                        </td>
                        <td />
                        <td>
                            <asp:DropDownList ID="ddlCountry" runat="server" Width="80%" Font-Size="XX-Small"
                                OnSelectedIndexChanged="ddlCountry_SelectedIndexChanged" AutoPostBack="True" />
                            &nbsp;<asp:LinkButton ID="lnkbtnUK" runat="server" OnClick="lnkbtnUK_Click" CausesValidation="False">UK</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trPostCode" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneePostCode" runat="server" ErrorMessage="###"
                                Font-Bold="True" Font-Names="Verdana" ControlToValidate="tbCneePostCode" />
                            &nbsp;<asp:Label ID="lblLegendPostcode" runat="server" Text="Postcode:" ForeColor="Red" />
                        </td>
                        <td align="left">
                        </td>
                        <td>
                            <asp:TextBox ID="tbCneePostCode" runat="server" Width="50%" MaxLength="50" Font-Size="XX-Small" />
                            &nbsp;<asp:Button ID="btnPostcodeFind" runat="server" Text="Find" OnClick="btnPostcodeFind_Click"
                                CausesValidation="False" />
                            <asp:Label ID="lblLookupError" runat="server" Visible="False" ForeColor="Red" Font-Size="XX-Small"
                                Font-Names="Verdana" Font-Bold="True"></asp:Label>
                        </td>
                    </tr>
                    <tr id="trPostcodeLookupOutput" runat="server" visible="false">
                        <td align="right">
                            &nbsp;<asp:Label ID="Label2" runat="server" Text="Select a destination:" />
                        </td>
                        <td />
                        <td>
                            <asp:ListBox ID="lbLookupResults" runat="server" Rows="10" Width="100%" OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged"
                                AutoPostBack="True" Font-Size="XX-Small"></asp:ListBox>
                            <br />
                            <asp:LinkButton ID="lnkbtnCancelPostcodeLookup" runat="server" OnClick="lnkbtnCancelPostcodeLookup_Click"
                                CausesValidation="False">cancel</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trContactName" runat="server">
                        <td align="right">
                            <asp:Label ID="Label4" runat="server" Text="Contact Name:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeCtcName" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCneeName" runat="server">
                        <td align="right" class="style2">
                            <asp:RequiredFieldValidator ID="rfvCneeName" runat="server" ErrorMessage="###" Font-Bold="True"
                                Font-Names="Verdana" ControlToValidate="tbCneeName" />
                            &nbsp;<asp:Label ID="Label5" runat="server" Text="Name:" ForeColor="Red" />
                        </td>
                        <td />
                        <td class="style2">
                            <asp:TextBox ID="tbCneeName" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr1" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneeAddr1" runat="server" ErrorMessage="###" Font-Bold="True"
                                Font-Names="Verdana" ControlToValidate="tbCneeAddr1"></asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="Label7" runat="server" Text="Addr 1:" ForeColor="Red" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr1" runat="server" Width="100%" Font-Size="XX-Small" MaxLength="50" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr2" runat="server">
                        <td align="right">
                            <asp:Label ID="Label9" runat="server" Text="Addr 2:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr2" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr3" runat="server">
                        <td align="right">
                            <asp:Label ID="Label10" runat="server" Text="Addr 3:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr3" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trTownCity" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvTownCity" runat="server" ErrorMessage="###" Font-Bold="True"
                                Font-Names="Verdana" ControlToValidate="tbCneeTown"></asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="Label11" runat="server" Text="Town/City:" ForeColor="Red" />
                        </td>
                        <td />
                        <td class="style4">
                            <asp:TextBox ID="tbCneeTown" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCneeState" runat="server">
                        <td align="right">
                            <asp:Label ID="lblLegendCountyStateRegionProvince" runat="server" Text="County / Region:" />
                        </td>
                        <td />
                        <td class="style4">
                            <asp:TextBox ID="tbCneeState" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trUSStatesCanadianProvinces" runat="server" visible="false">
                        <td align="right" class="style1">
                            &nbsp;<asp:CompareValidator ID="cvStateProvince" runat="server" ControlToValidate="ddlUSStatesCanadianProvinces"
                                Font-Names="Verdana" Operator="NotEqual" ValueToCompare="0" Text="###" Font-Size="XX-Small"
                                Font-Bold="True" />
                            &nbsp;<asp:Label ID="lblLegendStateProvince" runat="server" Text="State:" ForeColor="Red" />
                        </td>
                        <td />
                        <td class="style1">
                            <asp:DropDownList ID="ddlUSStatesCanadianProvinces" runat="server" Width="100%" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trUSStateShortcuts" runat="server" visible="false">
                        <td align="right">
                            &nbsp;
                        </td>
                        <td />
                        <td>
                            <asp:LinkButton ID="lnkbtnNewYorkCity" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                OnClick="lnkbtnNewYorkCity_Click" CausesValidation="False">NYC</asp:LinkButton>
                            &nbsp;<asp:LinkButton ID="lnkbtnWashingtonDC" runat="server" Font-Names="Verdana"
                                Font-Size="XX-Small" OnClick="lnkbtnWashingtonDC_Click" CausesValidation="False">Washington D.C.</asp:LinkButton>
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            &nbsp;
                        </td>
                        <td />
                        <td class="style4">
                            &nbsp;
                        </td>
                    </tr>
                    <tr id="trCneeTel" runat="server">
                        <td align="right">
                            <asp:Label ID="Label563" runat="server" Text="Contact Tel:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeTel" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCneeEmail" runat="server">
                        <td align="right">
                            <asp:Label ID="Label564" runat="server" Text="Contact Email:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeEmail" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCustRef1" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef1" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef1" Font-Names="Verdana"
                                Font-Size="XX-Small" Font-Bold="True">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef1" runat="server" Text="Cust Ref 1:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef1" runat="server" Width="100%" MaxLength="25" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCustRef2" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef2" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef2" Font-Names="Verdana"
                                Font-Size="XX-Small" Font-Bold="True">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef2" runat="server" Text="Cust Ref 2:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef2" runat="server" Width="100%" MaxLength="25" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCustRef3" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef3" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef3" Font-Names="Verdana"
                                Font-Size="XX-Small" Font-Bold="True">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef3" runat="server" Text="Cust Ref 3:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef3" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trCustRef4" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef4" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef4" Font-Names="Verdana"
                                Font-Size="XX-Small" Font-Bold="True">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef4" runat="server" Text="Cust Ref 4:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef4" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small" />
                        </td>
                    </tr>
                    <tr id="trArthritisCostCentres" runat="server" visible="false">
                        <td align="right">
                            <asp:RequiredFieldValidator
                                    ID="rfvArthritisCostCentre" runat="server" ControlToValidate="ddlArthritisCostCentre"
                                    ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" 
                                InitialValue="- please select -" Font-Bold="True" />
                        &nbsp;<asp:Label ID="lblLegendArthritisCostCentre" runat="server" 
                                Font-Size="XX-Small" Font-Names="Verdana"
                                Font-Bold="False" ForeColor="Red">Cost Centre:</asp:Label>
                        </td>
                        <td />
                        <td>
                            <asp:DropDownList ID="ddlArthritisCostCentre" runat="server" DataTextField="name"
                                DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsArthritisCostCentres"
                                ForeColor="Navy" TabIndex="12" Font-Names="Verdana" />
                        </td>
                    </tr>
                    <tr id="trArthritisCategories" runat="server" visible="false">
                        <td align="right">
                            <asp:RequiredFieldValidator
                                    ID="rfvArthritisCategory" runat="server" ControlToValidate="ddlArthritisCategory"
                                    ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" 
                                InitialValue="- please select -" Font-Bold="True" />
                        &nbsp;<asp:Label ID="Label135arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                                Font-Bold="False" ForeColor="Red">Category:</asp:Label>
                        </td>
                        <td />
                        <td>
                            <asp:DropDownList ID="ddlArthritisCategory" runat="server" DataTextField="name"
                                DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsArthritisCategories"
                                ForeColor="Navy" TabIndex="12" Font-Names="Verdana" />
                        </td>
                    </tr>
                    <tr id="trArthritisPONumber" runat="server" visible="false">
                        <td align="right">
                            <asp:RequiredFieldValidator
                                    ID="rfdArthritisPONumber" runat="server" ControlToValidate="tbArthritisPONumber"
                                    Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"> ###</asp:RequiredFieldValidator>
                        &nbsp;<asp:Label ID="lblLegendArthritisPONumber" runat="server" 
                                Font-Size="XX-Small" Font-Names="Verdana"
                                Font-Bold="False" ForeColor="Red">PO Number:</asp:Label>
                        </td>
                        <td />
                        <td>
                            <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                                ID="tbArthritisPONumber" Width="100%" MaxLength="25" />
                        </td>
                    </tr>
                    <tr id="trArthritisAdvisory" runat="server" visible="false">
                        <td align="right">

                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblLegendArthritisAdvisory" runat="server" 
                                Text="Type any further order instructions in Special Instructions" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            <asp:Label ID="Label18" runat="server" Text="Special Instrs:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbSpecialInstructions" runat="server" Width="100%" MaxLength="500"
                                Font-Size="XX-Small" Rows="3" TextMode="MultiLine" Font-Names="Verdana" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            <asp:Label ID="Label19" runat="server" Text="Packing Note:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbPackingNote" runat="server" Width="100%" MaxLength="50" Font-Size="XX-Small"
                                Rows="3" TextMode="MultiLine" Font-Names="Verdana" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <div id="divMainForm" runat="server">
        <table id="tblShowHideAddress" runat="server" visible="false">
            <tr>
                <td style="width: 110px" />
                <td style="width: 5px" />
                <td style="width: 300px">
                    <asp:LinkButton ID="lnkbtnShowHideAddress" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        OnClick="lnkbtnShowHideAddress_Click">hide address</asp:LinkButton>
                </td>
            </tr>
        </table>
    </div>
    <br />
    <br />
    <br />
    <br />
    <div id="divConfirmOrder" runat="server" style="background-color: Yellow; display: none;
        width: 300px">
        <br />
        <p style="text-align: center">
            <asp:Label ID="Label559" runat="server" Font-Names="Arial" Font-Size="Small" Text="Are you sure you want to submit this order?" />
            <br />
            <br />
            <asp:Button ID="btnOK" runat="server" Text="OK" Width="80px" OnClick="btnOK_Click"
                CausesValidation="false" />
            &nbsp;<asp:Button ID="btnCancel" runat="server" Text="Cancel" Width="80px" OnClick="btnCancel_Click"
                CausesValidation="false" />
        </p>
        <br />
    </div>
    <div id="divOrderError" runat="server" style="background-color: Yellow; display: none;
        width: 300px">
        <br />
        <p style="text-align: center">
            <asp:Label ID="Label8" runat="server" Font-Names="Arial" Font-Size="Small" Text="Sorry, we are unable to complete your order. Please contact Transworld Customer Services." />
            <br />
            <br />
            <asp:Button ID="btnContinueFromError" runat="server" Text="OK" Width="80px" CausesValidation="false" />
        </p>
        <br />
    </div>
    <br />
    <ajaxToolkit:ModalPopupExtender ID="mpe" TargetControlID="lnkbtnDummy" PopupControlID="divConfirmOrder"
        BackgroundCssClass="modalBackground" CancelControlID="btnCancel" runat="server" />
    <asp:LinkButton ID="lnkbtnDummy" runat="server" />
    <%--    <ajaxToolkit:TextBoxWatermarkExtender ID="TBWESearch" runat="server" TargetControlID="tbSearch"
        WatermarkText=" - search for products - " WatermarkCssClass="watermarked" />
--%>    <asp:XmlDataSource ID="xdsArthritisCostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
        XPath="OnLinePicksConfig/ArthritisCostCentres/costCentre" />
    <asp:XmlDataSource ID="xdsArthritisCategories" runat="server" DataFile="~/on_line_picks_config.xml"
        XPath="OnLinePicksConfig/ArthritisCategories/category" />
    <%--    <ajaxToolkit:TextBoxWatermarkExtender ID="TBWEQty" runat="server" TargetControlID="tbQty"
        WatermarkText=" " WatermarkCssClass="watermarked" />
    --%>
    </form>
</body>
</html>
