<%@ Page Language="VB" MasterPageFile="~/WebForm.master" Title="Online Ordering" StylesheetTheme="Basic" %>
<%@ MasterType virtualpath="~/WebForm.master" %>
<%@ Import Namespace=" System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>

<script runat="server">
    
    ' WebFormLogoLocation
    ' WebFormTopLegend
    ' WebFormBottomLegend
    ' WebFormPageTitle
    ' WebFormHomePageText
    ' WebFormAddressPageText
    ' WebFormHelpPageText
    ' WebFormShowPrice
    ' WebFormShowCostCentre

    'from SqlDataSourceAddressList <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
    ' hack with... <asp:Parameter DefaultValue="16" Name="CustomerKey" Type="Int32" />

    ' TO DO post UNIPUB release
    
    ' put base connection values in base web.config
    ' <%=GetPDFLink() %>    
    ' add keyboard navigation
    ' add client side button enabling / disabling depending on result of validation
    ' put CustomerKey in VIEWSTATE & use from there
    ' sort out duplicate stored procedures: ProductGetFromKey & GetProductFromKey (different param set)
    ' review product detail functionality esp. display of larger image
    ' allow for later addition of PDF downloads
    ' use of prod_image_folder, prod_thumb_folder
    ' integrate with web page editor
    ' disable Back to Product button on Basket panel if nothing in Products
    ' do something better than Server.Transfer("error.aspx") on db access error
    ' review all error handling
    ' rationalise sprocs

    ' SPROCS (last checked 18JUN08)
    'spASPNET_WebForm_GetCategories
    'spASPNET_WebForm_GetSubCategories
    'spASPNET_Product_GetFromKey5
    'spASPNET_ClientData_DrinkAware_Capture
    'spASPNET_Webform_GetProductFromKey
    'spASPNET_WebForm_GetTracking1a
    'spASPNET_WebForm_GetTracking2
    'spASPNET_WebForm_GetTracking3
    'spASPNET_Webform_AddBooking
    'spASPNET_LogisticMovement_Add
    'spASPNET_LogisticBooking_Complete

    'spASPNET_Webform_GetProductsUsingCategories
    'spASPNET_Webform_GetProductsUsingSearchCriteria
    'spASPNET_Country_GetCountries
    'spASPNET_Address_GetGlobalAddresses


    'Const ACCOUNT_CODE As String = "INDIV30495"
    Const ACCOUNT_CODE As String = "COURI11111"
    'Const LICENSE_KEY As String = "NA64-AU21-GN67-BR91"
    Const LICENSE_KEY As String = "RA61-XZ94-CT55-FH67"
    Dim gsConn As String = System.Configuration.ConfigurationManager.AppSettings("AIMSRootConnectionString")

    Private gdtBasket As DataTable = New DataTable()
    Private gdvBasketView As DataView
    Dim gnAffectedRows As Integer = 0
    Private gnZeroQuantityItems As Integer
    Private gsNotificationEmailAddr As String

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not pbWebformIsInitialised Then              ' is nothing on first load as panel has not yet been loaded to init SiteKey
            Call GetSiteKeyFromSiteName(sGetPath)
            If Not IsNumeric(Session("SiteKey")) Then
                WebMsgBox.Show("Could not initialise web form. Please inform Sprint Customer Services (customer_services@sprintexpress.co.uk). Thank you.")
                Exit Sub
            End If
            pbWebformIsInitialised = True
        End If
        If Not IsNumeric(Session("SiteKey")) Then
            Server.Transfer("timeout.aspx")
        End If
        If Not IsPostBack Then
            Dim sConfigPath As String = "~"
            Dim config As System.Configuration.Configuration = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(sConfigPath)
            Dim configSection As System.Web.Configuration.CompilationSection = _
              CType(config.GetSection("system.web/compilation"), _
                System.Web.Configuration.CompilationSection)
            Dim bDebug As Boolean = configSection.Debug
            If bDebug Then
                Call CheckSprocsExist()
            End If

            Call CheckSprocsExist()
            Call GetPageContent()

            Call GetCategories()
            Call ShowHome()

            lblBasketCount.Text = "0"
            Call AdjustBasketCountPlurality()
            pbCategoryProductsFound = False
            Session.Timeout = 180
            
            psVirtualThumbFolder = System.Configuration.ConfigurationManager.AppSettings("Virtual_Thumb_URL")
        End If

        tbPostCode.Attributes.Add("onkeypress", "return clickButton(event,'" + btnFindAddress.ClientID + "')")
        tbSearch.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGoSearch.ClientID + "')")
        tbConsignmentNo.Attributes.Add("onkeypress", "return clickButton(event,'" + btnCheckConsignment.ClientID + "')")
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
                    WebMsgBox.Show("This is web form is not fully configured. Path returned was '" & sSiteName & "'. Please inform your Account Handler.")
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
    
    Protected Sub GetPageContent()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_SiteContent2", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Action", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Action").Value = "GET"
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SiteKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@SiteKey").Value = Session("SiteKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ContentType", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ContentType").Value = "WebForm"

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("GetPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        Session("CustomerKey") = dr("WebFormCustomerKey")
        Session("GenericUserKey") = dr("WebFormGenericUserKey")
        Page.Header.Title = dr("WebFormPageTitle")

        Dim lbl As Label
        lbl = CType(Master.FindControl("lblTopLegend"), Label)
        If Not lbl Is Nothing Then
            lbl.Text = dr("WebFormTopLegend")
        End If
        lbl = CType(Master.FindControl("lblBottomLegend"), Label)
        If Not lbl Is Nothing Then
            lbl.Text = dr("WebFormBottomLegend")
        End If
        Dim img As Image
        img = CType(Master.FindControl("imgCompanyLogo"), Image)
        If Not img Is Nothing Then
            img.ImageUrl = dr("WebFormLogoImage")
        End If

        psHomePageText = dr("WebFormHomePageText")
        psAddressPageText = dr("WebFormAddressPageText")
        psHelpPageText = dr("WebFormHelpPageText")

        pbShowPrice = dr("WebFormShowPrice")
        pbShowZeroQuantity = dr("WebFormShowZeroQuantity")
        pbZeroStockNotification = dr("WebFormZeroStockNotification")
    End Sub

    Protected Function IsDAT() As Boolean
        IsDAT = (Session("CustomerKey") = 546)
    End Function
    
    Protected Function IsBRGIFTS() As Boolean
        IsBRGIFTS = (Session("CustomerKey") = 558)
    End Function
    
    Protected Function IsBlackRock() As Boolean
        IsBlackRock = (Session("CustomerKey") = 23)
    End Function
    
    Protected Sub PerCustomerSettings()
        If IsDAT() Then
            trCtcTelNo.Visible = True
            trCtcEmailAddr.Visible = True
            trTypeOfOrganisation.Visible = True
            trDrinkAwareOptIn.Visible = True
        End If
        If IsBRGIFTS() Or IsBlackRock() Then
            trCostCentre.Visible = True
        End If
    End Sub
    
    Protected Function OkayToChangePanels()
        If pnlBasket.Visible = True Then
            If RecordQuantities() = False Then
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If
    End Function
   
    Protected Sub HideAllPanels() ' ALL panel occlusion MUST come via here to ensure partially completed panel contents (eg Delivery Address) are saved
        pnlHome.Visible = False
        pnlCategorySelection.Visible = False
        pnlProductList.Visible = False
        pnlBasket.Visible = False
        pnlEmptyBasket.Visible = False
        pnlAddress.Visible = False
        pnlAddressList.Visible = False
        pnlSearch.Visible = False
        pnlSearchProductList.Visible = False
        pnlTrackAndTrace.Visible = False
        pnlTrackingResult.Visible = False
        pnlBookingConfirmation.Visible = False
        pnlHelp.Visible = False
        pnlRegisterOnly.Visible = False
        pnlRequestNotificationConfirmation.Visible = False
    End Sub
   
    Protected Sub ShowHome()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlHome.Visible = True
            lblHeading.Text = "Home"
            lblSubHeading.Text = psHomePageText
            lblBreadcrumbLocation.Text = "home"
        End If
    End Sub

    Protected Sub ShowCategories()
        Call HideAllPanels()
        pnlCategorySelection.Visible = True
        lblHeading.Text = "Product Categories"
        lblSubHeading.Text = "Choose a product category, then a sub-category."
        lblBreadcrumbLocation.Text = "products by category"
    End Sub

    Protected Sub ShowProductList()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlProductList.Visible = True
            gvCategoryProductList.PageIndex = 0
            Call BindCategoryProductList()
            lblHeading.Text = "Products"
            lblBreadcrumbLocation.Text = "products by category >>> products"
        End If
    End Sub
    
    Protected Sub BuildSubHeading(ByVal nRecordCount As Integer)
        Dim nProductCount As Integer = nRecordCount
        Dim sProductCountPluralisation As String = ""
        If nProductCount <> 1 Then
            sProductCountPluralisation = "s"
        End If
        If nProductCount > 0 Then
            pbCategoryProductsFound = True
        End If
        Dim sbSubHeadingText As StringBuilder = New StringBuilder()
        sbSubHeadingText.Append(nProductCount.ToString)
        sbSubHeadingText.Append(" product")
        sbSubHeadingText.Append(sProductCountPluralisation)
        lblSubHeading.Text = sbSubHeadingText.ToString
        sbSubHeadingText.Append(" found in category ")
        sbSubHeadingText.Append("<b>")
        sbSubHeadingText.Append(psCategory)
        sbSubHeadingText.Append("</b>")
        sbSubHeadingText.Append(", sub-category ")
        sbSubHeadingText.Append("<b>")
        sbSubHeadingText.Append(psSubCategory)
        sbSubHeadingText.Append("</b>")
        sbSubHeadingText.Append(". Click the <b>select</b> check box to choose your product, then click the <b>add to basket</b> button.")
        lblSubHeading.Text = sbSubHeadingText.ToString
    End Sub
   
    Protected Sub ShowSearch()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlSearch.Visible = True
            lblHeading.Text = "Search for a Product"
            lblSubHeading.Text = "Search the text of product descriptions."
            lblBreadcrumbLocation.Text = "product search"
            tbSearch.Focus()
        End If
    End Sub

    Protected Sub ShowSearchProductList()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlSearchProductList.Visible = True
            lblHeading.Text = "Search Results"
            Dim nSearchPageCount As Integer = gvSearchProductList.PageCount
            Dim sPage As String = "page"
            If nSearchPageCount > 1 Then sPage = sPage & "s"
            If gvSearchProductList.Rows.Count > 0 Then
                lblSubHeading.Text = "Your search term <b><i>""" & tbSearch.Text & """</i></b> matched the following products (" & nSearchPageCount.ToString & " " & sPage & " of results):"
                lnkbtnSearchAddToOrder.Visible = True
                lblBreadcrumbLocation.Text = "product search >>> search results"
                lnkbtnAddToOrder.Focus()
            Else
                lblSubHeading.Text = "No matching products found for your search term <b><i>""" & tbSearch.Text & """</i></b>"
                lblBreadcrumbLocation.Text = "product search >>> product list (none found)"
                lnkbtnSearchAddToOrder.Visible = False
            End If
        End If
    End Sub
   
    Protected Sub ShowCurrentBasket()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlBasket.Visible = True
            lblHeading.Text = "Your Basket"
            If CLng(lblBasketCount.Text) > 0 Then
                lblSubHeading.Text = "Indicate the quantity of each product you require, then click on Proceed to Checkout. You will be able to return here before finally submitting your order. At the next stage you will specify a delivery address, give any special instructions and provide any additional information required for this order."
            Else
                lblSubHeading.Text = ""
            End If
            lblBreadcrumbLocation.Text = "basket"
            If pbCategoryProductsFound = False Then
                lnkbtnBackToProducts.Visible = False
            Else
                lnkbtnBackToProducts.Visible = True
            End If
        End If
    End Sub
   
    Protected Sub ShowEmptyBasket()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlEmptyBasket.Visible = True
            lblHeading.Text = "Your Basket"
            lblSubHeading.Text = "Your basket is empty. Click on <b>products by category</b> to browse available items. Click on <b>product search</b> to search product descriptions."
            lblBreadcrumbLocation.Text = "basket (empty)"
        End If
    End Sub

    Protected Sub AdjustBasketCountPlurality()
        If CLng(lblBasketCount.Text) = 1 Then
            lblBasketCountPlural.Visible = False
        Else
            lblBasketCountPlural.Visible = True
        End If
    End Sub
   
    Protected Sub ShowAddressPanel()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlAddress.Visible = True
            lblHeading.Text = "Delivery Details"
            lblSubHeading.Text = psAddressPageText
            lblBreadcrumbLocation.Text = "basket >>> delivery details"
            tbPostCode.Focus()
        End If
    End Sub

    Protected Sub ShowAddressListPanel()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlAddressList.Visible = True
            tbSearchAddressList.Text = ""
            lblHeading.Text = "Address List"
            Call UpdateAddressListPanelLegends()
        End If
    End Sub
    
    Protected Sub UpdateAddressListPanelLegends()
        Dim nSearchPageCount As Integer = gvAddressList.PageCount
        Dim sPage As String = " page"
        If nSearchPageCount > 1 Then sPage = sPage & "s"
        If gvAddressList.Rows.Count > 0 Then
            gvAddressList.Visible = True
            lblSubHeading.Text = nSearchPageCount.ToString & sPage & " of addresses found.  Select a delivery address from the list by clicking on the Name.  Sort the columns by clicking on the column name.  Search for a specific name or part of a name using the Search box.  You can modify the address once selected."
            lblBreadcrumbLocation.Text = "delivery details >>> address list"
            tbSearchAddressList.Focus()
        Else
            gvAddressList.Visible = False
            lblSubHeading.Text = "No addresses found"
            lblBreadcrumbLocation.Text = "delivery details >>> address list (no addresses found)"
        End If
    End Sub

    Protected Sub ShowTrackAndTrace()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlTrackAndTrace.Visible = True
            lblHeading.Text = "Track And Trace"
            lblSubHeading.Text = "Enter the consignment number displayed when your order was placed."
            lblBreadcrumbLocation.Text = "track and trace"
            tbConsignmentNo.Focus()
        End If
    End Sub
   
    Protected Sub ShowTrackingResult()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlTrackingResult.Visible = True
            lblHeading.Text = "Tracking Result"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "track and trace >>> tracking result"
        End If
    End Sub
   
    Protected Sub ShowBookingConfirmation()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlBookingConfirmation.Visible = True
            lblHeading.Text = "Booking Confirmation"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "basket >>> delivery details >>> booking confirmation"
        End If
    End Sub
    
    Protected Sub ClearAddressPanel()
        tbTitle.Text = String.Empty
        tbName.Text = String.Empty
        tbAddr1.Text = String.Empty
        tbAddr2.Text = String.Empty
        tbTown.Text = String.Empty
        tbPostCode.Text = String.Empty
        tbAttnOf.Text = String.Empty
        tbCompany.Text = String.Empty
        tbCounty.Text = String.Empty
        tbCtcEmailAddr.Text = String.Empty
        tbCtcTelNo.Text = String.Empty
        tbJobTitle.Text = String.Empty
        tbSpclInstructions.Text = String.Empty
        Try
            ddlCountry.SelectedIndex = 0
            ddlCostCentre.SelectedIndex = 0
            ddlTypeOfOrganisation.SelectedIndex = 0
        Catch ex As Exception
            ' ddls in try block as they may not exist if only recording notification requests
        End Try
    End Sub
   
    Protected Sub ShowHelp()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlHelp.Visible = True
            lblHeading.Text = "Help"
            lblSubHeading.Text = psHelpPageText
            lblBreadcrumbLocation.Text = "help"
        End If
    End Sub
    
    Protected Sub ShowRegisterOnlyPanel()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlRegisterOnly.Visible = True
            lblHeading.Text = "Register for product availability email notifications"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "register"
        End If
    End Sub
  
    Protected Sub ShowRequestNotificationConfirmation()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlRequestNotificationConfirmation.Visible = True
            lblHeading.Text = "Registration confirmation"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "registration confirmation"
        End If
    End Sub
    
    Protected Sub lnkbtn_ShowProductsByCategory_click(ByVal sender As Object, ByVal e As CommandEventArgs)   ' user clicked on a sub-category so show list of products
        psSubCategory = CStr(e.CommandArgument)
        ShowProductList()
    End Sub

    Protected Sub rptrCategories_Item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = System.Drawing.Color.FromArgb(131, 148, 140)
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = System.Drawing.Color.FromArgb(242, 173, 13) 'selected
        lblSubCategoryHeading.Visible = True
    End Sub

    Protected Sub lnkbtn_ShowSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psCategory = CStr(e.CommandArgument)
        rptrSubCategories.Visible = True
        GetSubCategories()
    End Sub

    Protected Sub DisplayCategories()
        Dim item As RepeaterItem
        For Each item In rptrCategories.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = System.Drawing.Color.FromArgb(131, 148, 140)
        Next
        rptrSubCategories.Visible = False
        lblSubCategoryHeading.Visible = False
        Call ShowCategories()
    End Sub

    Protected Sub GetCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_WebForm_GetCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GenericUserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@GenericUserKey").Value = Session("GenericUserKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ShowZeroQuantity", SqlDbType.Bit))
        If pbShowZeroQuantity Then
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 1
        Else
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 0
        End If
        
        Try
            oAdapter.Fill(oDataSet, "Categories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                rptrCategories.DataSource = oDataSet
                rptrCategories.DataBind()
            End If
        Catch ex As SqlException
            Server.Transfer("error.aspx")
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub GetSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_WebForm_GetSubCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GenericUserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@GenericUserKey").Value = Session("GenericUserKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ShowZeroQuantity", SqlDbType.Bit))
        If pbShowZeroQuantity Then
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 1
        Else
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 0
        End If

        Try
            oAdapter.Fill(oDataSet, "SubCategories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                rptrSubCategories.Visible = True
                rptrSubCategories.DataSource = oDataSet
                rptrSubCategories.DataBind()
            Else
                rptrSubCategories.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub AddItemToBasket(ByVal sProductKey As String)
        Dim dr As DataRow
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetFromKey5", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int)
        oParamProductKey.Value = CLng(sProductKey)

        Dim oParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParamUserKey.Value = Session("GenericUserKey")

        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
           
            If IsNothing(Session("BasketData")) Then            ' create a new Basket structure & assign it to session variable
                gdtBasket = New DataTable()
                gdtBasket.Columns.Add(New DataColumn("ProductKey", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("ProductCode", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("ProductDate", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("Description", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("BoxQty", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("UnitWeightGrams", GetType(Double)))
                gdtBasket.Columns.Add(New DataColumn("UnitValue", GetType(Double)))
                gdtBasket.Columns.Add(New DataColumn("QtyAvailable", GetType(Long)))
                gdtBasket.Columns.Add(New DataColumn("QtyToPick", GetType(Long)))
                gdtBasket.Columns.Add(New DataColumn("QtyRequested", GetType(Long)))
                gdtBasket.Columns.Add(New DataColumn("PDFFileName", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("OriginalImage", GetType(String)))
                gdtBasket.Columns.Add(New DataColumn("ThumbNailImage", GetType(String)))
                Session("BasketData") = gdtBasket
            End If

            gdtBasket = Session("BasketData")
            gdvBasketView = New DataView(gdtBasket)
            gdvBasketView.RowFilter = "ProductKey='" & sProductKey & "'"   ' is selected product already in Basket?
            If gdvBasketView.Count = 0 Then                                ' no, so add it to Basket
                dr = gdtBasket.NewRow()
                dr("ProductKey") = sProductKey
                If Not IsDBNull(oDataReader("ProductCode")) Then
                    dr("ProductCode") = oDataReader("ProductCode")
                End If
                If Not IsDBNull(oDataReader("ProductDate")) Then
                    dr("ProductDate") = oDataReader("ProductDate")
                End If
                If Not IsDBNull(oDataReader("ProductDescription")) Then
                    dr("Description") = oDataReader("ProductDescription")
                End If
                If Not IsDBNull(oDataReader("ItemsPerBox")) Then
                    dr("BoxQty") = oDataReader("ItemsPerBox")
                End If
                If Not IsDBNull(oDataReader("UnitWeightGrams")) Then
                    dr("UnitWeightGrams") = oDataReader("UnitWeightGrams")
                End If
                If Not IsDBNull(oDataReader("UnitValue")) Then
                    dr("UnitValue") = oDataReader("UnitValue")
                End If
                If Not IsDBNull(oDataReader("Quantity")) Then
                    dr("QtyAvailable") = oDataReader("Quantity")
                Else
                    dr("QtyAvailable") = 0
                End If
                'If Not IsDBNull(oDataReader("MaxGrab")) Then
                ' dr("QtyToPick") = oDataReader("MaxGrab")
                'Else
                dr("QtyToPick") = 0
                'End If
                If Not IsDBNull(oDataReader("PDFFileName")) Then
                    dr("PDFFileName") = oDataReader("PDFFileName")
                End If
                If Not IsDBNull(oDataReader("OriginalImage")) Then
                    dr("OriginalImage") = oDataReader("OriginalImage")
                End If
                If Not IsDBNull(oDataReader("ThumbNailImage")) Then
                    dr("ThumbNailImage") = oDataReader("ThumbNailImage")
                End If
                gdtBasket.Rows.Add(dr)
                lblBasketCount.Text = CLng(lblBasketCount.Text) + 1             ' increment Basket item count
                Call AdjustBasketCountPlurality()
                Session("BasketData") = gdtBasket
                gdvBasketView.RowFilter = ""
            End If
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub RemoveItemFromBasket(ByVal sProductKey As String)
        gdtBasket = Session("BasketData")
        gdvBasketView = New DataView(gdtBasket)
        gdvBasketView.RowFilter = "ProductKey='" & sProductKey & "'"               ' set filter to selected record
        If gdvBasketView.Count > 0 Then                                            ' if record is present
            gdvBasketView.Delete(0)                                                ' remove it
            lblBasketCount.Text = CLng(lblBasketCount.Text) - 1
            Call AdjustBasketCountPlurality()
        End If
        gdvBasketView.RowFilter = ""
        Session("BasketData") = gdtBasket
    End Sub
   
    Protected Sub btnRemoveItemFromBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call btnRemoveItemFromBasket_Click()
    End Sub
   
    Protected Sub btnRemoveItemFromBasket_Click()                               ' NB: each item in Basket has a Remove checkbox and associated hidden field that holds the Product Key
        Dim cbRemoveItem As CheckBox
        Dim hidProductKey As HiddenField
        For Each row As GridViewRow In gvBasket.Rows
            cbRemoveItem = row.FindControl("cbRemoveItemFromBasket")
            If cbRemoveItem.Checked = True Then
                hidProductKey = row.FindControl("hidProductKey")
                RemoveItemFromBasket(hidProductKey.Value)
            End If
        Next
        Call BindBasketGrid("ProductCode")
    End Sub

    Protected Sub BindBasketGrid(ByVal SortField As String)
        If Not IsNothing(Session("BasketData")) Then
            gdtBasket = Session("BasketData")
            gdvBasketView = New DataView(gdtBasket)
            gdvBasketView.Sort = SortField
            If gdvBasketView.Count > 0 Then
                gvBasket.DataSource = gdvBasketView
                gvBasket.DataBind()
                If gnZeroQuantityItems > 0 Then
                    Dim sStockAvailabilityNotificationMessage As String
                    If gnZeroQuantityItems = 1 Then
                        lblZeroQuantity.Text = "Please note: we are unable to deliver the item highlighted in red in your basket as it is currently out of stock. At the checkout stage you can register to receive an email notification when this item becomes available again."
                        sStockAvailabilityNotificationMessage = "Enter your email address below if you would like to be notified<br /> when the out of stock item becomes available again."
                        cbRegister.Text = "Yes, send me an email when the out of stock item becomes available"
                    Else
                        lblZeroQuantity.Text = "Please note: we are unable to deliver the items highlighted in red in your basket as they are currently out of stock. At the checkout stage you can register to receive an email notification when these items become available again."
                        sStockAvailabilityNotificationMessage = "Enter your email address below if you would like to be notified<br /> when the out of stock itema become available again."
                        cbRegister.Text = "Yes, send me an email when the out of stock items become available"
                    End If
                    lblStockAvailabilityNotificationMessage1.Text = sStockAvailabilityNotificationMessage
                    lblStockAvailabilityNotificationMessage2.Text = sStockAvailabilityNotificationMessage
                    tblZeroQuantity.Visible = True
                Else
                    tblZeroQuantity.Visible = False
                End If
                ShowCurrentBasket()
            Else
                ShowEmptyBasket()
            End If
        Else
            ShowEmptyBasket()
        End If
    End Sub

    Protected Sub lnkbtnHome_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowHome()
    End Sub

    Protected Sub lnkbtnProductsByCategory_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllAddressFields()
        lnkbtnSubmitOrder.Visible = True
        Call ShowCategories()
    End Sub

    Protected Sub lnkbtnProductSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowSearch()
    End Sub

    Protected Sub lnkbtnMyOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCurrentBasket()
        BindBasketGrid("ProductCode")
    End Sub

    Protected Sub lnkbtnTrackAndTrace_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowTrackAndTrace()
    End Sub

    Protected Sub lnkbtnHelpWithOrdering_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowHelp()
    End Sub

    Protected Sub lnkbtnAddToOrderFromCategories_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bItemAdded As Boolean = False
        Dim cbAddToOrder As CheckBox
        Dim hidProductKey As HiddenField
        For Each row As GridViewRow In gvCategoryProductList.Rows
            cbAddToOrder = row.FindControl("cbAddToOrder")
            If cbAddToOrder.Checked = True Then
                hidProductKey = row.FindControl("hidProductKey")
                AddItemToBasket(hidProductKey.Value)
                cbAddToOrder.Checked = False
                bItemAdded = True
            End If
        Next
        If bItemAdded Then              ' only show basket if something was added to it
            Call BindBasketGrid("ProductCode")
            Call ShowCurrentBasket()
        End If
    End Sub

    Protected Sub lnkbtnAddToOrderFromSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bItemAdded As Boolean = False
        Dim cbAddToOrder As CheckBox
        Dim hidProductKey As HiddenField
        For Each row As GridViewRow In gvSearchProductList.Rows
            cbAddToOrder = row.FindControl("cbAddToOrder")
            If cbAddToOrder.Checked = True Then
                hidProductKey = row.FindControl("hidProductKey")
                AddItemToBasket(hidProductKey.Value)
                cbAddToOrder.Checked = False
                bItemAdded = True
            End If
        Next
        If bItemAdded Then              ' only show basket if something was added to it
            Call BindBasketGrid("ProductCode")
            Call ShowCurrentBasket()
        End If
    End Sub

    Protected Sub btnSubmitOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nConsignmentKey As Integer
        Page.Validate("Address")
        If Page.IsValid Then
            If cbRegister.Checked Then
                gsNotificationEmailAddr = tbNotificationEmailAddr1.Text
                Call RecordNotificationAddresses()
            End If
            nConsignmentKey = SubmitOrder()
            If nConsignmentKey > 0 Then
                If IsDAT() Then
                    Call CaptureData(nConsignmentKey)
                End If
                Call ShowBookingConfirmation()
                Call TidyUp()
                Call ClearAddressPanel()
                If gnZeroQuantityItems > 0 Then
                    Dim sItem As String
                    If gnZeroQuantityItems = 1 Then
                        sItem = "out of stock item becomes"
                    Else
                        sItem = gnZeroQuantityItems.ToString & " out of stock items become"
                    End If
                    WebMsgBox.Show("Thank you for your order. We will notify you when the " & sItem & "available.")
                End If
            Else
                '
            End If
        End If
    End Sub

    Protected Sub SearchProductDescriptions()
        gvSearchProductList.PageIndex = 0
        Call BindSearchProductList()
        Call ShowSearchProductList()
    End Sub
    
    Protected Function GetPDFLink() As String
        Dim sVirtualPDFFolder As String = System.Configuration.ConfigurationManager.AppSettings("Virtual_PDF_URL")
        Dim sProdPDFFolder As String = System.Configuration.ConfigurationManager.AppSettings("prod_pdf_folder")
        Dim sbHTML As StringBuilder = New StringBuilder()
        Dim t As String = "                                        " 'tabs for indenting :)
   
        If File.Exists(MapPath(sVirtualPDFFolder & Session("ProductKey") & ".pdf")) Then
            sbHTML.Append(t & "<br />" & vbCrLf)
            sbHTML.Append(t & "<br />" & vbCrLf)
            sbHTML.Append(t & "<img src=" & sVirtualPDFFolder & "pdf_logo_small.gif>&nbsp;<a href=" & Chr(34) & sVirtualPDFFolder & Session("ProductKey") & ".pdf" & Chr(34) & " Target='_blank'>Click here to download an electronic version of this product.</a>" & vbCrLf)
        End If
        Return sbHTML.ToString()
    End Function
   
    Protected Sub CheckConsignment()
        Page.Validate("tracking")
        If Page.IsValid Then
            Call ShowTrackingResult()
        End If
    End Sub
    
    Protected Function GetTracking() As String
        GetTracking = String.Empty
        If Not IsNumeric(Session("SiteKey")) Then
            Exit Function   ' reqd because this is called automatically from HTML and SiteKey may be undefined
        End If
        Dim sbHTML As StringBuilder = New StringBuilder()
        Dim t As String = "                                        " 'tabs for indenting
        Dim sConsignee As String = ""
        Dim sNOP As String = ""
        Dim sWeight As String = ""
        Dim sPODStatus As String = ""
        Dim lConsignmentKey As Long
        Dim oDataReader1 As SqlDataReader = Nothing
        Dim bRecordFound As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_WebForm_GetTracking1a", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim oParam1 As New SqlParameter("@ConsignmentNo", SqlDbType.NVarChar, 50)
        oCmd.Parameters.Add(oParam1)
        oParam1.Value = tbConsignmentNo.Text

        Dim oParam2 As New SqlParameter("@CustomerKey", SqlDbType.Int)
        oCmd.Parameters.Add(oParam2)
        oParam2.Value = Session("CustomerKey")

        Try
            oConn.Open()
            oDataReader1 = oCmd.ExecuteReader()
            lblTrackingMessage.Text = ""
            While oDataReader1.Read()
                bRecordFound = True
                If Not IsDBNull(oDataReader1("Key")) Then
                    lConsignmentKey = CLng(oDataReader1("Key"))
                End If
                If Not IsDBNull(oDataReader1("CneeName")) Then
                    If oDataReader1("CneeName") <> "" Then
                        sConsignee = oDataReader1("CneeName") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeAddr1")) Then
                    If oDataReader1("CneeAddr1") <> "" Then
                        sConsignee &= oDataReader1("CneeAddr1") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeAddr2")) Then
                    If oDataReader1("CneeAddr2") <> "" Then
                        sConsignee &= oDataReader1("CneeAddr2") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeAddr3")) Then
                    If oDataReader1("CneeAddr3") <> "" Then
                        sConsignee &= oDataReader1("CneeAddr3") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeTown")) Then
                    If oDataReader1("CneeTown") <> "" Then
                        sConsignee &= oDataReader1("CneeTown") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeState")) Then
                    If oDataReader1("CneeState") <> "" Then
                        sConsignee &= oDataReader1("CneeState") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneePostCode")) Then
                    If oDataReader1("CneePostCode") <> "" Then
                        sConsignee &= oDataReader1("CneePostCode") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CountryName")) Then
                    If oDataReader1("CountryName") <> "" Then
                        sConsignee &= oDataReader1("CountryName") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeCtcName")) Then
                    If oDataReader1("CneeCtcName") <> "" Then
                        sConsignee &= oDataReader1("CneeCtcName") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("CneeTel")) Then
                    If oDataReader1("CneeTel") <> "" Then
                        sConsignee &= oDataReader1("CneeTel") & "<br>"
                    End If
                End If
                If Not IsDBNull(oDataReader1("NOP")) Then
                    sNOP = oDataReader1("NOP")
                Else
                    sNOP = "0"
                End If
                If Not IsDBNull(oDataReader1("Weight")) Then
                    sWeight = oDataReader1("Weight")
                Else
                    sWeight = "0"
                End If
                If Not IsDBNull(oDataReader1("PODName")) Then
                    If oDataReader1("PODName") <> "" Then
                        sPODStatus = " Status: DELIVERED [" & oDataReader1("PODName") & " " & oDataReader1("PODDate") & " " & oDataReader1("PODTime") & "]<br>"
                    End If
                Else
                    sPODStatus &= " Status: ON_ROUTE" & "<br>"
                End If
                sbHTML.Append(vbCrLf)
                sbHTML.Append(t & "Consignment number " & tbConsignmentNo.Text & "<br /><b>" & sPODStatus & "</b>" & vbCrLf)
                sbHTML.Append("<br />")
                sbHTML.Append(t & sConsignee & vbCrLf)
            End While
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oDataReader1.Close()
        End Try
   
        If bRecordFound Then
            Dim oDataSet1 As New DataSet()
            Dim oDataTable1 As DataTable
            Dim oAdapter1 As New SqlDataAdapter("spASPNET_WebForm_GetTracking2", oConn)
            oAdapter1.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter1.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oAdapter1.SelectCommand.Parameters("@ConsignmentKey").Value = lConsignmentKey
            Try
                oAdapter1.Fill(oDataSet1, "Tracking1")
                oDataTable1 = oDataSet1.Tables(0)
                Dim iCount As Integer = oDataTable1.Rows.Count
                If iCount > 0 Then
   
                    sbHTML.Append(vbCrLf)
                    sbHTML.Append(t & "<table class='listing' width='95%'>" & vbCrLf)
                    sbHTML.Append(t & " <thead>" & vbCrLf)
                    sbHTML.Append(t & "  <tr>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Product Code</th>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Description</th>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Quantity</th>" & vbCrLf)
                    sbHTML.Append(t & "  </tr>" & vbCrLf)
                    sbHTML.Append(t & " </thead>" & vbCrLf)
                    sbHTML.Append(t & " <tbody>" & vbCrLf)
   
                    Dim oRow As DataRow
                    For Each oRow In oDataTable1.Rows
                        Dim sProductCode As String = oRow.Item("ProductCode").ToString()
                        Dim sDescription As String = oRow.Item("ProductDescription").ToString()
                        Dim sQuantity As String = oRow.Item("ItemsOut").ToString()
                        sbHTML.Append(vbCrLf)
                        sbHTML.Append(t & "<tr>" & vbCrLf)
                        sbHTML.Append(t & " <td style='text-align: center;'>" & vbCrLf)
                        sbHTML.Append(t & sProductCode & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & " <td>" & vbCrLf)
                        sbHTML.Append(t & sDescription & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & " <td style='text-align: right;'>" & vbCrLf)
                        sbHTML.Append(t & sQuantity & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & "</tr>" & vbCrLf)
                    Next
                    sbHTML.Append(t & " </tbody>" & vbCrLf)
                    sbHTML.Append(t & "</table>" & vbCrLf)
                    sbHTML.Append(vbCrLf)
                Else
                End If
            Catch ex As SqlException
                Server.Transfer(" error.aspx")
            End Try
   
            Dim oDataSet2 As New DataSet()
            Dim oDataTable2 As DataTable
            Dim oAdapter2 As New SqlDataAdapter("spASPNET_WebForm_GetTracking3", oConn)
            oAdapter2.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter2.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oAdapter2.SelectCommand.Parameters("@ConsignmentKey").Value = lConsignmentKey
            Try
                oAdapter2.Fill(oDataSet2, "Tracking2")
                oDataTable2 = oDataSet2.Tables(0)
                Dim iCount As Integer = oDataTable2.Rows.Count
                If iCount > 0 Then
                    sbHTML.Append(vbCrLf)
                    sbHTML.Append(t & "<table class='listing' width='95%'>" & vbCrLf)
                    sbHTML.Append(t & " <thead>" & vbCrLf)
                    sbHTML.Append(t & "  <tr>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Tracked On</th>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Location</th>" & vbCrLf)
                    sbHTML.Append(t & "   <th>Tracking Event</th>" & vbCrLf)
                    sbHTML.Append(t & "  </tr>" & vbCrLf)
                    sbHTML.Append(t & " </thead>" & vbCrLf)
                    sbHTML.Append(t & " <tbody>" & vbCrLf)
                    Dim oRow As DataRow
                    For Each oRow In oDataTable2.Rows
                        Dim sTrackedOn As String = oRow.Item("Time").ToString()
                        Dim sLocation As String = oRow.Item("Location").ToString()
                        Dim sTrackingEvent As String = oRow.Item("Description").ToString()
                        sbHTML.Append(vbCrLf)
                        sbHTML.Append(t & "<tr>" & vbCrLf)
                        sbHTML.Append(t & " <td>" & vbCrLf)
                        sbHTML.Append(t & sTrackedOn & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & " <td>" & vbCrLf)
                        sbHTML.Append(t & sLocation & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & " <td>" & vbCrLf)
                        sbHTML.Append(t & sTrackingEvent & vbCrLf)
                        sbHTML.Append(t & " </td>" & vbCrLf)
                        sbHTML.Append(t & "</tr>" & vbCrLf)
                    Next
                    sbHTML.Append(t & " </tbody>" & vbCrLf)
                    sbHTML.Append(t & "</table>" & vbCrLf)
                    sbHTML.Append(vbCrLf)
                Else
                    '
                End If
            Catch ex As SqlException
                Server.Transfer("error.aspx")
            Finally
                oConn.Close()
            End Try
            Call ShowTrackingResult()
        Else
            lblTrackingMessage.Text = "Consignment not found."
        End If
        Return sbHTML.ToString()
    End Function

    Protected Function SubmitOrder() As Integer
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim bBookingFailed As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_Webform_AddBooking2", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = Session("GenericUserKey")
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = Session("CustomerKey")
        oCmdAddBooking.Parameters.Add(param2)
        Dim param3 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param3.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param3)
        Dim param4 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param4.Value = Nothing
        oCmdAddBooking.Parameters.Add(param4)
        Dim param5 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param5.Value = "PRINTED MATTER - FREE DOMICILE"
        oCmdAddBooking.Parameters.Add(param5)
        Dim param6 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param6.Value = tbName.Text
        oCmdAddBooking.Parameters.Add(param6)
        Dim param7 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param7.Value = tbCompany.Text
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param8.Value = tbAddr1.Text
        oCmdAddBooking.Parameters.Add(param8)
        Dim param9 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param9.Value = tbAddr2.Text
        oCmdAddBooking.Parameters.Add(param9)
        Dim param10 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param10.Value = tbTown.Text
        oCmdAddBooking.Parameters.Add(param10)
        Dim param11 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param11.Value = tbCounty.Text
        oCmdAddBooking.Parameters.Add(param11)
        Dim param12 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param12.Value = tbPostCode.Text
        oCmdAddBooking.Parameters.Add(param12)
        Dim param13 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param13.Value = CLng(ddlCountry.SelectedValue)
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param14.Value = tbAttnOf.Text
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param15.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param16.Value = tbSpclInstructions.Text
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        If IsBRGIFTS() Or IsBlackRock() Then
            Dim sTemp As String = ddlCostCentre.SelectedItem.Text
            If sTemp.Length > 25 Then
                sTemp = sTemp.Substring(sTemp.Length - 25).Trim
            End If
            param17.Value = sTemp
        Else
            param17.Value = String.Empty
        End If
        oCmdAddBooking.Parameters.Add(param17)
   
        Dim param21 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param21.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param22.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param22)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            'Output parameter contains the new Booking Key
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value)
            If lBookingKey > 0 Then
                gdtBasket = Session("BasketData")
                gdvBasketView = gdtBasket.DefaultView
                Dim ProductItem As DataRowView
                If gdvBasketView.Count > 0 Then
                    For Each ProductItem In gdvBasketView
                        If ProductItem("QtyAvailable") > 0 Then
                            Dim lProductKey As Long = CLng(ProductItem("ProductKey"))
                            Dim lPickQuantity As Long = CLng(ProductItem("QtyRequested"))
                            Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                            oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                            Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                            param51.Value = Session("GenericUserKey")
                            oCmdAddStockItem.Parameters.Add(param51)
                            Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                            param52.Value = Session("CustomerKey")
                            oCmdAddStockItem.Parameters.Add(param52)
                            Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                            param53.Value = lBookingKey
                            oCmdAddStockItem.Parameters.Add(param53)
                            Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                            param54.Value = lProductKey
                            oCmdAddStockItem.Parameters.Add(param54)
                            Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                            param55.Value = "RECEIVED"
                            oCmdAddStockItem.Parameters.Add(param55)
                            Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                            param56.Value = lPickQuantity
                            oCmdAddStockItem.Parameters.Add(param56)
                            Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
                            param57.Value = lConsignmentKey
                            oCmdAddStockItem.Parameters.Add(param57)
                            oCmdAddStockItem.Connection = oConn
                            oCmdAddStockItem.Transaction = oTrans
                            oCmdAddStockItem.ExecuteNonQuery()
                        End If
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
                    bBookingFailed = True
                    lblError.Text = "No stock items found for booking"
                End If
            Else
                bBookingFailed = True
                lblError.Text = "Error adding Web Booking [BookingKey=0]."
            End If
            If Not bBookingFailed Then
                oTrans.Commit()
                lblConsignmentKey.Text = lConsignmentKey
            Else
                oTrans.Rollback("AddBooking")
            End If
            Session.Clear()
        Catch ex As SqlException
            'Server.Transfer("error.aspx")
            lblError.Text = ex.ToString
            oTrans.Rollback("AddBooking")   ' WORK OUT WHY COMPILER IS WARNING ABOUT THIS STATEMENT
        Finally
            oConn.Close()
        End Try
        If bBookingFailed Then
            Return 0
        Else
            Return lConsignmentKey
        End If
    End Function
   
    Protected Sub lnkbtnBackToProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllAddressFields()
        lnkbtnSubmitOrder.Visible = True
        Call ShowProductList()
    End Sub

    Protected Sub lnkbtnProceedToCheckout_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("Basket")
        If Page.IsValid Then
            gdtBasket = Session("BasketData")
            Dim nZeroQuantityItem As Integer
            Dim nNonZeroQuantityItem As Integer
            For Each dr As DataRow In gdtBasket.Rows
                If dr("QtyAvailable") = 0 Then
                    nZeroQuantityItem += 1
                Else
                    nNonZeroQuantityItem += 1
                End If
            Next
            If nZeroQuantityItem > 0 Then
                tblRegister.Visible = True
            End If
            If nNonZeroQuantityItem > 0 Then
                Call ShowAddressPanel()
                'If nNonZeroQuantityItem = 1 Then
                ' cbRegister.Text = "Yes, send me an email when the out of stock item becomes available"
                'Else
                '   cbRegister.Text = "Yes, send me an email when the out of stock items become available"
                'End If
            Else
                tbNotificationEmailAddr2.Focus()
                Call ShowRegisterOnlyPanel()
            End If
        End If
    End Sub

    Protected Sub lnkbtnRemoveItemFromBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call btnRemoveItemFromBasket_Click()
    End Sub
   
    Function RecordQuantities() As Boolean
        Dim bQuantitiesOkay As Boolean = True
        Dim gvr As GridViewRow
        Dim tbQuantity As TextBox
        Dim hidProductKey As HiddenField
        Dim lblProductCode As Label
        Dim hidQtyAvailable As HiddenField
        Dim x As String
        gdtBasket = Session("BasketData")
        gdvBasketView = New DataView(gdtBasket)

        For Each gvr In gvBasket.Rows
            hidProductKey = gvr.FindControl("hidProductKey")
            hidQtyAvailable = gvr.FindControl("hidQtyAvailable")
            lblProductCode = gvr.FindControl("lblProductCode")
            tbQuantity = gvr.FindControl("tbQuantity")
            x = tbQuantity.Text
            gdvBasketView.RowFilter = "ProductKey='" & hidProductKey.Value & "'"
            If gdvBasketView.Count = 1 AndAlso IsNumeric(tbQuantity.Text) Then
                If CLng(hidQtyAvailable.Value) >= CLng(tbQuantity.Text) Then
                    gdvBasketView(0).Item("QtyRequested") = CLng(tbQuantity.Text)
                Else
                    lblInsufficientQuantityAvailable.Text = "Insufficent quantity available for product " & lblProductCode.Text & "!"
                    lblInsufficientQuantityAvailable.Visible = True
                    bQuantitiesOkay = False
                    Exit For
                End If
            End If
        Next
        Return bQuantitiesOkay
    End Function
   
    Protected Sub TidyUp()
        Session.Clear()
        pbCategoryProductsFound = False
        lblBasketCount.Text = "0"
        Call AdjustBasketCountPlurality()
        Call GetSiteKeyFromSiteName(sGetPath)
        Call GetPageContent()
    End Sub

    Protected Sub gvProductList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim cbAddToOrder As CheckBox
        Dim gvrea As GridViewRowEventArgs = e
        If gvrea.Row.RowType = DataControlRowType.DataRow Then
            Dim nQtyAvailable As Integer = CInt(DataBinder.Eval(e.Row.DataItem, "QtyAvailable"))
            If nQtyAvailable = 0 And Not pbZeroStockNotification Then
                cbAddToOrder = e.Row.FindControl("cbAddToOrder")
                cbAddToOrder.Enabled = False
            End If
        End If
        If Not pbShowPrice Then
            If gvrea.Row.Cells.Count >= 3 Then
                gvrea.Row.Cells(3).Visible = False
            End If
        End If
    End Sub
    
    Protected Sub btnAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAddressListPanel()
        gvAddressList.DataBind()
    End Sub

    Protected Sub lnkbtnBackToDeliveryAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAddressPanel()
    End Sub

    Protected Sub lnkbtnSelectedAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtnSelAddr As LinkButton
        Dim hid As HiddenField, gvr As GridViewRow
        
        lnkbtnSelAddr = CType(sender, LinkButton)
        gvr = lnkbtnSelAddr.NamingContainer
        Dim sSpclInstructions As String
        Dim nCountryCode As Integer, sCountryName As String

        hid = CType(gvr.FindControl("hidCountryCode"), HiddenField)
        nCountryCode = CInt(hid.Value)
        
        hid = CType(gvr.FindControl("hidDefaultSpecialInstructions"), HiddenField)
        sSpclInstructions = hid.Value

        tbAttnOf.Text = CType(gvr.FindControl("lblAttnOf"), Label).Text
        tbName.Text = lnkbtnSelAddr.Text
        tbAddr1.Text = CType(gvr.FindControl("lblAddr1"), Label).Text
        tbAddr2.Text = CType(gvr.FindControl("lblAddr2"), Label).Text
        'tbAddr3.Text = CType(gvr.FindControl("lblAddr3"), Label).Text
        tbTown.Text = CType(gvr.FindControl("lblTown"), Label).Text
        tbCounty.Text = CType(gvr.FindControl("lblState"), Label).Text
        tbPostCode.Text = CType(gvr.FindControl("lblPostCode"), Label).Text
        'tbCtcTelNo.Text = CType(gvr.FindControl("lblTelephone"), Label).Text
        If Len(sSpclInstructions) > 0 Then
            tbSpclInstructions.Text = sSpclInstructions
        End If
        sCountryName = CType(gvr.FindControl("lblCountryName"), Label).Text
        
        Dim s As ListItem
        Dim counter As Integer = 0
        For Each s In ddlCountry.Items
            counter += 1
            If s.Text = sCountryName Then
                ddlCountry.SelectedIndex = counter - 1
                Exit For
            End If
        Next
        Call ShowAddressPanel()
    End Sub

    Protected Sub lnkbtnSearchAddressList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvAddressList.DataBind()
        Call UpdateAddressListPanelLegends()
    End Sub

    Protected Sub FindAddress()
        lnkbtnSubmitOrder.Visible = False
        lblAddressValidation.Visible = False
        lblLookupError.Text = String.Empty
        tbPostCode.Text = tbPostCode.Text.Trim.ToUpper
        trAddr1.Visible = False
        trAddr2.Visible = False
        trAddr2a.Visible = False
        trAddr3.Visible = False
        trAddr4.Visible = False
        trAddr6.Visible = False
        trAddr7.Visible = False
        trAddr9.Visible = False
        trCostCentre.Visible = False
        trAddr10.Visible = False
        
        trTitle.Visible = False
        trJobTitle.Visible = False
        trCtcTelNo.Visible = False
        trCtcEmailAddr.Visible = False
        trTypeOfOrganisation.Visible = False
        trDrinkAwareOptIn.Visible = False

        trPostcodeLookupResults.Visible = True

        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult

        objInterimResults = objLookup.ByPostcode(tbPostCode.Text, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        
        If objInterimResults.IsError OrElse objInterimResults.Results Is Nothing OrElse objInterimResults.Results.GetLength(0) = 0 Then
            lblLookupError.Visible = True
            lbLookupResults.Visible = False
            lblSelectADestination.Visible = False
            lblLookupError.Text = objInterimResults.ErrorMessage
            If lblLookupError.Text.Trim = String.Empty Then
                lblLookupError.Text = "<br />No results found for this post code"
            Else
                lblLookupError.Text = "<br />" & lblLookupError.Text
            End If
        Else
            lblLookupError.Visible = False
            lbLookupResults.Visible = True
            lblSelectADestination.Visible = True
            lbLookupResults.Items.Clear()
            If Not objInterimResults.Results Is Nothing Then      ' add the new items to the list
                For Each objInterimResult In objInterimResults.Results
                    lbLookupResults.Items.Add(New _
                         ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            Dim oCmd As SqlCommand
            Dim oConn As New SqlConnection(gsConn)
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            Dim sSQL As String = "INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('"
            sSQL += sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 0)"
            Try
                oConn.Open()
                oCmd = New SqlCommand(sSQL, oConn)
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                NotifyException("lnkbtnFindAddress_Click", "Could not log lookup", ex)
            Finally
                oConn.Close()
            End Try
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
        
        Call ShowAllAddressFields()
        lnkbtnSubmitOrder.Visible = True
        lblAddressValidation.Visible = True
        tbAttnOf.Focus()

        If objAddressResults.IsError Then
            lblLookupError.Text = objAddressResults.ErrorMessage
        Else
            objAddress = objAddressResults.Results(0)

            tbCompany.Text = objAddress.OrganisationName
            tbAddr1.Text = objAddress.Line1
            tbAddr2.Text = objAddress.Line2
            'tbAddr3.Text = objAddress.Line3
            tbTown.Text = objAddress.PostTown
            tbPostCode.Text = objAddress.Postcode
            tbCounty.Text = objAddress.County

            For i As Integer = 0 To ddlCountry.Items.Count
                If ddlCountry.Items(i).Text = "U.K." Then
                    ddlCountry.SelectedIndex = i
                    Exit For
                End If
            Next
            Dim oCmd As SqlCommand
            Dim oConn As New SqlConnection(gsConn)
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            Dim sSQL As String = "INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('"
            sSQL += sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)"
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            Try
                oCmd.ExecuteNonQuery()
            Catch ex As Exception
                NotifyException("lbLookupResults_SelectedIndexChanged", "Could not log lookup", ex)
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub lnkbtnAddrLookupCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblLookupError.Text = String.Empty
        lnkbtnSubmitOrder.Visible = True
        lblAddressValidation.Visible = True
        Call ShowAllAddressFields()
    End Sub
    
    Protected Sub ShowAllAddressFields()
        trAddr1.Visible = True
        trAddr2.Visible = True
        trAddr2a.Visible = True
        trAddr3.Visible = True
        trAddr4.Visible = True
        trAddr6.Visible = True
        trAddr7.Visible = True
        trAddr9.Visible = True
        If IsBRGIFTS() Or IsBlackRock() Then
            trCostCentre.Visible = True
        End If
        trAddr10.Visible = True
        If IsDAT() Then
            trTitle.Visible = True
            trJobTitle.Visible = True
            trCtcTelNo.Visible = True
            trCtcEmailAddr.Visible = True
            trTypeOfOrganisation.Visible = True
            trDrinkAwareOptIn.Visible = True
        End If
        trPostcodeLookupResults.Visible = False
        lblLookupError.Text = String.Empty
        lblAddressValidation.Visible = True
    End Sub
    
    Protected Sub CaptureData(ByVal nConsignmentKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_ClientData_DrinkAware_Capture", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim param1 As SqlParameter = New SqlParameter("@Title", SqlDbType.VarChar, 50)
        param1.Value = tbTitle.Text.Trim
        oCmd.Parameters.Add(param1)
        
        Dim param2 As SqlParameter = New SqlParameter("@Name", SqlDbType.VarChar, 80)
        param2.Value = tbName.Text.Trim
        oCmd.Parameters.Add(param2)
        
        Dim param3 As SqlParameter = New SqlParameter("@Company", SqlDbType.VarChar, 80)
        param3.Value = tbCompany.Text.Trim
        oCmd.Parameters.Add(param3)
        
        Dim param4 As SqlParameter = New SqlParameter("@JobTitle", SqlDbType.VarChar, 80)
        param4.Value = tbJobTitle.Text.Trim
        oCmd.Parameters.Add(param4)
        
        Dim sAddress As String = String.Empty
        Dim arrAddress As New ArrayList
        
        Dim sTemp As String

        sTemp = tbAddr1.Text.Trim
        If sTemp.Length > 0 Then
            arrAddress.Add(sTemp)
        End If
        
        sTemp = tbAddr2.Text.Trim
        If sTemp.Length > 0 Then
            arrAddress.Add(sTemp)
        End If
        
        sTemp = tbTown.Text.Trim
        If sTemp.Length > 0 Then
            arrAddress.Add(sTemp)
        End If
        
        sTemp = tbCounty.Text.Trim
        If sTemp.Length > 0 Then
            arrAddress.Add(sTemp)
        End If

        If arrAddress.Count >= 2 Then
            For i As Integer = 0 To arrAddress.Count - 2
                sAddress = sAddress & arrAddress(i) & ", "
            Next
            sAddress = sAddress & arrAddress(arrAddress.Count - 1)
        Else
            sAddress = arrAddress(0)
        End If
        
        Dim param5 As SqlParameter = New SqlParameter("@Address", SqlDbType.VarChar, 250)
        param5.Value = sAddress
        oCmd.Parameters.Add(param5)
        
        Dim param6 As SqlParameter = New SqlParameter("@Postcode", SqlDbType.VarChar, 50)
        param6.Value = tbPostCode.Text.Trim
        oCmd.Parameters.Add(param6)
        
        Dim param7 As SqlParameter = New SqlParameter("@Email", SqlDbType.VarChar, 80)
        param7.Value = tbCtcEmailAddr.Text.Trim
        oCmd.Parameters.Add(param7)
        
        Dim param8 As SqlParameter = New SqlParameter("@Telephone", SqlDbType.VarChar, 50)
        param8.Value = tbCtcTelNo.Text.Trim
        oCmd.Parameters.Add(param8)
        
        Dim param9 As SqlParameter = New SqlParameter("@Comments", SqlDbType.VarChar, 1000)
        param9.Value = tbSpclInstructions.Text.Trim
        oCmd.Parameters.Add(param9)
        
        Dim param10 As SqlParameter = New SqlParameter("@TypeOfOrganisation", SqlDbType.VarChar, 50)
        param10.Value = ddlTypeOfOrganisation.SelectedValue
        oCmd.Parameters.Add(param10)
        
        Dim param11 As SqlParameter = New SqlParameter("@OptIn", SqlDbType.Bit)
        param11.Value = cbOptIn.Checked
        oCmd.Parameters.Add(param11)
        
        Dim param12 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int)
        param12.Value = nConsignmentKey
        oCmd.Parameters.Add(param12)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch
            
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub NotifyException(ByVal sLocation As String, ByVal sReason As String, Optional ByVal ex As Exception = Nothing, Optional ByVal bContinue As Boolean = False, Optional ByVal sAdviceString As String = "")
        Dim sbMessage As New StringBuilder
        sbMessage.Append(sReason & " in " & sLocation)
        If ex IsNot Nothing Then
            sbMessage.Append(vbCrLf & vbCrLf & "Exception: ")
            sbMessage.Append(ex.Message & vbCrLf & vbCrLf)
            sbMessage.Append("Stack Trace: ")
            sbMessage.Append(ex.StackTrace & vbCrLf & vbCrLf)
        End If
        If sAdviceString.Length > 0 Then
            sbMessage.Append(sAdviceString)
        End If
        WebMsgBox.Show(sbMessage.ToString.Replace("'", "*").Replace("""", "*").Replace(vbLf, "").Replace(vbCr, "\n"))
    End Sub
    
    Protected Sub lnkbtnCountryUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Integer = 0 To ddlCountry.Items.Count - 1
            If ddlCountry.Items(i).Text = "U.K." Then
                ddlCountry.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Protected Sub btnFindAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call FindAddress()
    End Sub
    
    Protected Sub btnGoSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SearchProductDescriptions()
    End Sub
    
    Protected Sub btnCheckConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CheckConsignment()
    End Sub
    
    Protected Sub SqlDataSourceCategoryProductList_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs)
        If Len(ViewState("WF_SubCategory")) > 0 Then
            e.Command.Parameters("@Category").Value = ViewState("WF_Category")
            e.Command.Parameters("@SubCategory").Value = ViewState("WF_SubCategory")
        End If
    End Sub

    Protected Sub SqlDataSourceCategoryProductList_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs)
        Dim sdssea As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs = e
        gnAffectedRows = e.AffectedRows
    End Sub
    
    Protected Sub SqlDataSourceSearchProductList_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs)
        Dim sdssea As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs = e
        gnAffectedRows = e.AffectedRows
    End Sub

    Protected Sub BindCategoryProductList()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Webform_GetProductsUsingCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GenericUserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@GenericUserKey").Value = Session("GenericUserKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ShowZeroQuantity", SqlDbType.Bit))
        If pbShowZeroQuantity Then
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 1
        Else
            oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = 0
        End If

        Try
            oAdapter.Fill(oDataTable)
            gvCategoryProductList.DataSource = oDataTable
            gvCategoryProductList.DataBind()
            Call BuildSubHeading(oDataTable.Rows.Count)
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub BindSearchProductList()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Webform_GetProductsUsingSearchCriteria2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = tbSearch.Text

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GenericUserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@GenericUserKey").Value = Session("GenericUserKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ShowZeroQuantity", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@ShowZeroQuantity").Value = pbShowZeroQuantity

        Try
            oAdapter.Fill(oDataTable)
            gvSearchProductList.DataSource = oDataTable
            gvSearchProductList.DataBind()
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub CheckSprocsExist()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("select name from sysobjects where name like 'spaspnet_webform%'", oConn)
        Dim dt As New DataTable

        Dim sSprocNames As New ArrayList
        sSprocNames.Add("spASPNET_WebForm_GetCategories")
        sSprocNames.Add("spASPNET_Product_GetSubCategories")
        sSprocNames.Add("spASPNET_WebForm_GetCategories2")
        sSprocNames.Add("spASPNET_Product_GetSubCategories2")
        sSprocNames.Add("spASPNET_Product_GetFromKey5")          ' used in AddItemToBasket, only uses ProductKey
        sSprocNames.Add("spASPNET_Webform_GetProductFromKey")   ' used in GetProductDetail, takes CustomerKey to do EXECUTE spWAddLogisticWebHit
        sSprocNames.Add("spASPNET_WebForm_GetTracking1a")
        sSprocNames.Add("spASPNET_WebForm_GetTracking2")
        sSprocNames.Add("spASPNET_WebForm_GetTracking3")
        sSprocNames.Add("spASPNET_Webform_AddBooking")
        sSprocNames.Add("spASPNET_LogisticMovement_Add")
        sSprocNames.Add("spASPNET_LogisticBooking_Complete")

        sSprocNames.Add("spASPNET_Webform_GetProductsUsingCategories")
        sSprocNames.Add("spASPNET_Webform_GetProductsUsingSearchCriteria")
        sSprocNames.Add("spASPNET_Webform_GetProductsUsingCategories2")
        sSprocNames.Add("spASPNET_Webform_GetProductsUsingSearchCriteria2")
        sSprocNames.Add("spASPNET_Country_GetCountries")
        Dim s As String
        For Each s In sSprocNames
            Try
                oAdapter.SelectCommand.CommandText = "SELECT name FROM sysobjects WHERE name LIKE '%" & s & "%'"
                dt.Clear()
                oAdapter.Fill(dt)
                If dt.Rows.Count = 0 Then
                    Throw New SystemException(s & " not defined")
                End If
            Catch ex As Exception
                WebMsgBox.Show(ex.ToString)
            End Try
        Next
    End Sub
    
    Protected Sub gvCategoryProductList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvCategoryProductList.PageIndex = e.NewPageIndex
        Call BindCategoryProductList()
    End Sub

    Protected Sub gvSearchProductList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvSearchProductList.PageIndex = e.NewPageIndex
        Call BindSearchProductList()
    End Sub
    
    Protected Sub gvBasket_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gv As GridView = sender
        Dim gvr As GridViewRow = e.Row
        Dim hidQtyAvailable As HiddenField
        Dim tb As TextBox
        Dim rfv As RequiredFieldValidator
        If gvr.RowType = DataControlRowType.DataRow Then
            hidQtyAvailable = gvr.FindControl("hidQtyAvailable")
            tb = gvr.FindControl("tbQuantity")
            rfv = gvr.FindControl("rfvQuantity")
            If hidQtyAvailable.Value = 0 Then
                ' gvr.BackColor = Red
                'gvr.Font.Strikeout = True
                gvr.ForeColor = Red
                'gvr.CssClass = "RedBold"
                tb.Visible = False
                rfv.Enabled = False
                gnZeroQuantityItems += 1
            End If
        End If
    End Sub
    
    Protected Sub cbRegister1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            tbNotificationEmailAddr1.Enabled = True
            tbConfirmNotificationEmailAddr1.Enabled = True
            
            revNotificationEmailAddr1.Enabled = True
            revNotificationEmailAddr1.EnableClientScript = True

            revConfirmNotificationEmailAddr1.Enabled = True
            revConfirmNotificationEmailAddr1.EnableClientScript = True
            
            rfvNotificationEmailAddr1.Enabled = True
            rfvNotificationEmailAddr1.EnableClientScript = True

            rfvConfirmNotificationEmailAddr1.Enabled = True
            rfvConfirmNotificationEmailAddr1.EnableClientScript = True
            
            cvNotificationEmailAddr1.Enabled = True
            cvNotificationEmailAddr1.EnableClientScript = True
        Else
            tbNotificationEmailAddr1.Enabled = False
            tbConfirmNotificationEmailAddr1.Enabled = False
            
            revNotificationEmailAddr1.Enabled = False
            revNotificationEmailAddr1.EnableClientScript = False

            revConfirmNotificationEmailAddr1.Enabled = False
            revConfirmNotificationEmailAddr1.EnableClientScript = False
            
            rfvNotificationEmailAddr1.Enabled = False
            rfvNotificationEmailAddr1.EnableClientScript = False

            rfvConfirmNotificationEmailAddr1.Enabled = False
            rfvConfirmNotificationEmailAddr1.EnableClientScript = False
            
            cvNotificationEmailAddr1.Enabled = False
            cvNotificationEmailAddr1.EnableClientScript = False
        End If
    End Sub
    
    Protected Sub lnkbtnSubmitRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("NotificationEmailAddr2")
        If Page.IsValid Then
            gsNotificationEmailAddr = tbNotificationEmailAddr2.Text
            Call RecordNotificationAddresses()
            Dim sItem As String
            If gnZeroQuantityItems = 1 Then
                sItem = "out of stock item is"
            Else
                sItem = gnZeroQuantityItems.ToString & " out of stock items are"
            End If
            lblRequestNotificationConfirmation.Text = "Thank you. We will send you an email when the " & sItem & " available."
            Call ShowRequestNotificationConfirmation()
            Call TidyUp()
            tbNotificationEmailAddr2.Text = String.Empty
            tbConfirmNotificationEmailAddr2.Text = String.Empty
            Call ClearAddressPanel()
        End If
    End Sub
    
    Protected Sub RecordNotificationAddresses()
        gdtBasket = Session("BasketData")
        For Each dr As DataRow In gdtBasket.Rows
            If dr("QtyAvailable") = 0 Then
                Call RecordNotificationAddress(dr("ProductKey"), gsNotificationEmailAddr.Trim, 0)
                gnZeroQuantityItems += 1
            End If
        Next
    End Sub

    Protected Sub RecordNotificationAddress(ByVal nLogisticProductKey As Integer, ByVal sEmailAddr As String, ByVal nQuantityRequired As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_ZeroStockNotification_Record", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = nLogisticProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = 0
        oCmd.Parameters.Add(paramUserKey)
        
        Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.VarChar, 100)
        paramEmailAddr.Value = sEmailAddr
        oCmd.Parameters.Add(paramEmailAddr)
        
        Dim paramQuantityRequired As SqlParameter = New SqlParameter("@QuantityRequired", SqlDbType.Int)
        paramQuantityRequired.Value = nQuantityRequired
        oCmd.Parameters.Add(paramQuantityRequired)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in RecordNotificationAddress: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Property pbShowPrice() As Boolean
        Get
            Dim o As Object = ViewState("WF_ShowPrice")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WF_ShowPrice") = Value
        End Set
    End Property
    
    Property pbShowZeroQuantity() As Boolean
        Get
            Dim o As Object = ViewState("WF_ShowZeroQuantity")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WF_ShowZeroQuantity") = Value
        End Set
    End Property
    
    Property pbZeroStockNotification() As Boolean
        Get
            Dim o As Object = ViewState("WF_ZeroStockNotification")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WF_ZeroStockNotification") = Value
        End Set
    End Property
    
    Property psVirtualThumbFolder() As String
        Get
            Dim o As Object = ViewState("WF_VirtualThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_VirtualThumbFolder") = Value
        End Set
    End Property

    Property psHomePageText() As String
        Get
            Dim o As Object = ViewState("WF_HomePageText")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_HomePageText") = Value
        End Set
    End Property

    Property psAddressPageText() As String
        Get
            Dim o As Object = ViewState("WF_AddressPageText")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_AddressPageText") = Value
        End Set
    End Property

    Property psHelpPageText() As String
        Get
            Dim o As Object = ViewState("WF_HelpPageText")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_HelpPageText") = Value
        End Set
    End Property

    Property psCategory() As String
        Get
            Dim o As Object = ViewState("WF_Category")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_Category") = Value
        End Set
    End Property

    Property psSubCategory() As String
        Get
            Dim o As Object = ViewState("WF_SubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("WF_SubCategory") = Value
        End Set
    End Property

    Property pbCategoryProductsFound() As Boolean
        Get
            Dim o As Object = ViewState("WF_CategoryProductsFound")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WF_CategoryProductsFound") = Value
        End Set
    End Property

    Property pbWebformIsInitialised() As Boolean
        Get
            Dim o As Object = ViewState("WF_WebformIsInitialised")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("WF_WebformIsInitialised") = Value
        End Set
    End Property

</script>

<asp:Content ID="ContentBreadcrumbs" ContentPlaceHolderID="ContentPlaceHolderBreadcrumbs" runat="Server">
    <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
        <tr>
            <td style="white-space: nowrap; width:40%">
                &nbsp;you are here:
                <asp:Label ID="lblBreadcrumbLocation" runat="server" Text="home"></asp:Label></td>
            <td style="white-space: nowrap; width:60%">
                &nbsp; - your basket contains
                <asp:Label runat="server" ID="lblBasketCount">0</asp:Label>
                item<asp:Label runat="server" ID="lblBasketCountPlural">s</asp:Label>
                -
            </td>
        </tr>
    </table>
</asp:Content>
<asp:Content ID="ContentNavigation" ContentPlaceHolderID="ContentPlaceHolderNavigation"
    runat="Server">
    <table border="0" cellspacing="0" cellpadding="0" id="navigation" width="185px">
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnHome" runat="server" Width="135px" OnClick="lnkbtnHome_Click"> &nbsp;home&nbsp;</asp:LinkButton></a>
            </td>
        </tr>
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnProductsByCategory" runat="server" Width="135px" OnClick="lnkbtnProductsByCategory_click"> &nbsp;products&nbsp;by&nbsp;category&nbsp;</asp:LinkButton></a>
            </td>
        </tr>
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnProductSearch" runat="server" Width="135px" OnClick="lnkbtnProductSearch_Click"> &nbsp;product&nbsp;search&nbsp;</asp:LinkButton></a></td>
        </tr>
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnMyOrder" runat="server" Width="135px" OnClick="lnkbtnMyOrder_Click"> &nbsp;your&nbsp;basket&nbsp;</asp:LinkButton></a>
            </td>
        </tr>
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnTrackAndTrace" runat="server" Width="135px" OnClick="lnkbtnTrackAndTrace_Click"> &nbsp;track&nbsp;&&nbsp;trace&nbsp;</asp:LinkButton></a>
            </td>
        </tr>
        <tr>
            <td style="width: 135px">
                <a href="" class="navText">
                    <asp:LinkButton ID="lnkbtnHelpWithOrdering" runat="server" Width="135px" OnClick="lnkbtnHelpWithOrdering_Click"> &nbsp;help&nbsp;</asp:LinkButton></a>
            </td>
        </tr>
    </table>
</asp:Content>
<asp:Content ID="ContentMain" ContentPlaceHolderID="ContentPlaceHolderMain" runat="Server">
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
            <td class="pageName" style="height: 20px; width: 856px;">
                <asp:Label ID="lblHeading" runat="server" Text=""/>
            </td>
        </tr>
        <tr>
            <td class="bodyText" style="height: 15px; width: 856px;">
                <br />
                <p>
                    <asp:Label ID="lblSubHeading" runat="server" />&nbsp;</p>
            </td>
        </tr>
        <tr>
            <td class="bodyText" style="width: 856px">
                <asp:Panel ID="pnlHome" runat="server" Visible="False" Width="100%">
                    <table style="width:100%">
                        <tr>
                            <td align="center" style="white-space: nowrap">
                                <br />
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                            <br />
                            <br />
                            <br />
                            <br />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlCategorySelection" runat="server" Visible="True" Width="100%">
                    <table style="width:100%">
                        <tr>
                            <td style="width: 5%"></td>
                            <td style="width: 30%; white-space: nowrap">
                                <br />
                            </td>
                            <td style="width: 65%"></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td valign="top" style="white-space: nowrap">
                                <asp:Label ID="Label14" runat="server" ForeColor="#72837b" Font-Bold="True">&nbsp;&nbsp; Category</asp:Label>
                                <br />
                                <br />
                                <asp:Repeater runat="server" ID="rptrCategories" OnItemCommand="rptrCategories_Item_click">
                                    <ItemTemplate>
                                        <asp:Image ID="Image1" runat="server" ImageUrl="./images/greycircle.gif"></asp:Image>
                                        <asp:LinkButton ID="lnkbtnShowSubCategories" runat="server" OnCommand="lnkbtn_ShowSubCategories_click"
                                            CommandArgument='<%# Container.DataItem("ProductCategory")%>' Text='<%# Container.DataItem("ProductCategory")%>'
                                            ForeColor="#83948C" />
                                        <br />
                                        <br />
                                    </ItemTemplate>
                                </asp:Repeater>
                                <br />
                            </td>
                            <td valign="top" style="white-space: nowrap">
                                <asp:Label runat="server" ID="lblSubCategoryHeading" ForeColor="#72837b" Font-Bold="True"
                                    Visible="False">&nbsp;&nbsp; Sub-Category</asp:Label>
                                <br />
                                <br />
                                <asp:Repeater runat="server" Visible="False" ID="rptrSubCategories">
                                    <ItemTemplate>
                                        <asp:Image ID="Image2" runat="server" ImageUrl="./images/greycircle.gif" />
                                        <asp:LinkButton ID="lnkbtnShowProductsByCategory" runat="server" OnCommand="lnkbtn_ShowProductsByCategory_click"
                                            CommandArgument='<%# Container.DataItem("SubCategory")%>' Text='<%# Container.DataItem("SubCategory")%>'
                                            ForeColor="#83948C" />
                                        <br />
                                        <br />
                                    </ItemTemplate>
                                </asp:Repeater>
                                <br />
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlProductList" runat="server" Visible="False" Width="100%">
                    <br />
                    <asp:GridView ID="gvCategoryProductList" runat="server" AllowPaging="True" AutoGenerateColumns="False" PageSize="5" GridLines="None" PagerSettings-Mode="NextPreviousFirstLast" OnRowDataBound="gvProductList_RowDataBound" OnPageIndexChanging="gvCategoryProductList_PageIndexChanging">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>' />
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Product Code">
                                <ItemTemplate>
                                    <asp:HiddenField ID="hidProductKey" Value='<%# Bind("LogisticProductKey") %>' runat="server" />
                                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("ProductCode") %>' Width="89px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Height="50px" BorderColor="#DDDDDD" BorderStyle="Dotted"
                                    BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description" SortExpression="ProductDescription">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="Label1" runat="server" Text='<%# Bind("ProductDescription") %>'
                                        Width="400px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Price" SortExpression="UnitValue">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="Label1a" runat="server" Text='<%# String.Format("{0:c}", Eval("UnitValue")) %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="&nbsp;Qty Available&nbsp;&nbsp;" SortExpression="QtyAvailable">
                                <ItemTemplate>
                                    <asp:Label ID="Label3" runat="server" Text='<%# Bind("MaxGrab") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="&nbsp;&nbsp;select&nbsp;">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbAddToOrder" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                        </Columns>
                        <PagerSettings Mode="NumericFirstLast" />
                        <PagerStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:GridView>
                    <br />
                    <table style="width:100%">
                        <tr valign="middle">
                            <td style="width:10%" >&nbsp;</td>
                            <td style="width:30%" >&nbsp;</td>
                            <td style="width:60%" valign="middle" align="center">
                                <asp:LinkButton ID="lnkbtnAddToOrder" runat="server" OnClick="lnkbtnAddToOrderFromCategories_Click"
                                    SkinID="button">&nbsp;add&nbsp;to&nbsp;basket&nbsp;</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlSearchProductList" runat="server" Visible="False" Width="100%">
                    <br />
                    <asp:GridView ID="gvSearchProductList" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                        PageSize="5" GridLines="None" PagerSettings-Mode="NextPreviousFirstLast" OnRowDataBound="gvProductList_RowDataBound" OnPageIndexChanging="gvSearchProductList_PageIndexChanging">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>' />
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Product Code">
                                <ItemTemplate>
                                    <asp:HiddenField ID="hidProductKey" Value='<%# Bind("LogisticProductKey") %>' runat="server" />
                                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("ProductCode") %>' Width="89px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" Height="50px" BorderColor="#DDDDDD" BorderStyle="Dotted"
                                    BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description" SortExpression="ProductDescription">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="Label1" runat="server" Text='<%# Bind("ProductDescription") %>'
                                        Width="400px"></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Price" SortExpression="UnitValue">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="Label1aa" runat="server" Text='<%# String.Format("{0:c}", Eval("UnitValue")) %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="&nbsp;Qty Available&nbsp;&nbsp;" SortExpression="QtyAvailable">
                                <ItemTemplate>
                                    <asp:Label ID="Label3a" runat="server" Text='<%# Bind("MaxGrab") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="&nbsp;&nbsp;select&nbsp;">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbAddToOrder" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                        </Columns>
                        <PagerSettings Mode="NumericFirstLast" />
                        <PagerStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    </asp:GridView>
                    <br />
                    <table style="width:100%">
                        <tr valign="middle">
                            <td style="width: 10%"></td>
                            <td style="width: 30%"></td>
                            <td style="width: 60%" valign="middle" align="center">
                                <asp:LinkButton ID="lnkbtnSearchAddToOrder" runat="server" OnClick="lnkbtnAddToOrderFromSearch_Click"
                                    SkinID="button">&nbsp;add&nbsp;to&nbsp;basket&nbsp;</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlBasket" runat="server" Visible="False" Width="100%">
                    <table style="width:100%">
                        <tr>
                            <td style="width: 10%">&nbsp;</td>
                            <td style="width: 80%"></td>
                            <td style="width: 10%"></td>
                        </tr>
                    </table>
                    <asp:GridView ID="gvBasket" runat="server" AutoGenerateColumns="False" Width="100%"
                        ShowFooter="True" GridLines="None" OnRowDataBound="gvBasket_RowDataBound">
                        <Columns>
                            <asp:TemplateField HeaderText="remove item">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbRemoveItemFromBasket" runat="server" />
                                    <asp:HiddenField ID="hidProductKey" Value='<%# Bind("ProductKey") %>' runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <FooterTemplate>
                                    <img src="./images/1TxparentPixel.gif" alt="" height="20px" width="1px" />
                                    <asp:LinkButton ID="lnkbtnRemoveItemFromBasket" runat="server" OnClick="lnkbtnRemoveItemFromBasket_Click"
                                        SkinID="button">&nbsp;remove&nbsp;</asp:LinkButton>
                                </FooterTemplate>
                                <FooterStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Product Code">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="lblProductCode" runat="server" Text='<%# Bind("ProductCode") %>'
                                        Width="90px" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Description">
                                <ItemTemplate>
                                    <asp:Label ID="lblProductDescription" runat="server" Text='<%# Bind("Description") %>'
                                        Width="400px" />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Price" SortExpression="UnitValue">
                                <ItemTemplate>
                                    &nbsp;<asp:Label ID="Label1aaa" runat="server" Text='<%# String.Format("{0:c}", Eval("UnitValue")) %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                                <HeaderStyle HorizontalAlign="Center" Width="100px" Wrap="False" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Quantity">
                                <ItemTemplate>
                                    <itemstyle width="45px">
                                <asp:TextBox id="tbQuantity" runat="server" Text='<% # Bind("QtyRequested") %>' Width="30px" MaxLength="4"/><asp:RequiredFieldValidator id="rfvQuantity" runat="server" ValidationGroup="Basket" ErrorMessage="&nbsp;required!&nbsp;" ControlToValidate="tbQuantity" Display="Dynamic"/><asp:RangeValidator id="rvQuantity" runat="server" ValidationGroup="Basket" ErrorMessage="invalid quantity!" ControlToValidate="tbQuantity" Display="Dynamic" MaximumValue="99999" MinimumValue="1"></asp:RangeValidator> <asp:HiddenField id="hidQtyAvailable" runat="server" Value='<%# Bind("QtyAvailable") %>'/>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                    <table style="width: 100%" id="tblZeroQuantity" runat="server" visible="false">
                        <tr>
                            <td style="width: 10%">
                            </td>
                            <td style="width: 80%">
                                <asp:Label ID="lblZeroQuantity" Font-Names="Verdana" Font-Size="Small" runat="server" Font-Bold="True" ForeColor="Red"/>
                            </td>
                            <td style="width: 10%">
                            </td>
                        </tr>
                    </table>
                    <table style="width:100%">
                        <tr valign="middle">
                            <td style="width:10%"></td>
                            <td style="width:30%">
                                <asp:Label ID="lblInsufficientQuantityAvailable" runat="server" Text="Insufficient quantity available!"
                                    ForeColor="Red" Visible="false" EnableViewState="false"/>
                            </td>
                            <td style="width:10%"></td>
                            <td valign="middle" style="width:60%" >
                                <asp:LinkButton ID="lnkbtnBackToCategories" runat="server" OnClick="lnkbtnProductsByCategory_Click"
                                    CausesValidation="false" SkinID="button">&nbsp;back&nbsp;to&nbsp;categories&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnBackToProducts" runat="server" OnClick="lnkbtnBackToProducts_Click"
                                    CausesValidation="false" SkinID="button">&nbsp;back&nbsp;to&nbsp;products&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnProceedToCheckout" runat="server" OnClick="lnkbtnProceedToCheckout_Click"
                                    SkinID="button">&nbsp;proceed&nbsp;to&nbsp;checkout&nbsp;</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlEmptyBasket" runat="server" Visible="False">
                </asp:Panel>
                <asp:Panel ID="pnlAddress" runat="server" Width="100%">
                    <br />
                    <table width="90%">
                        <tr>
                            <td style="width: 20%">
                                <strong>Post Code</strong>
                            </td>
                            <td style="width: 65%">
                                <asp:TextBox ID="tbPostCode" runat="server" Width="400px" MaxLength="50"></asp:TextBox>&nbsp;
                                <asp:Button ID="btnFindAddress" runat="server" Text="find address" OnClick="btnFindAddress_Click" />&nbsp;
                                <asp:Label ID="lblLookupError" runat="server" Visible="False" ForeColor="Red"/>
                            </td>
                            <td align="left" style="width: 15%">
                            </td>
                        </tr>
                        <tr id="trPostcodeLookupResults" runat="server" visible="false">
                            <td>
                                <asp:Label ID="lblSelectADestination" runat="server" Text="Select a destination"></asp:Label></td>
                            <td>
                                <br />
                                <asp:ListBox ID="lbLookupResults" runat="server" AutoPostBack="True" Width="408px"
                                    OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged" Height="250px"></asp:ListBox>
                                <asp:LinkButton ID="lnkbtnAddrLookupCancel" runat="server" OnClick="lnkbtnAddrLookupCancel_Click">cancel</asp:LinkButton></td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trAddr1" runat="server" visible="true">
                            <td>
                                Attention of: <span style="font-size: xx-small">(for use when you are not the initial
                                    recipient)</span>
                            </td>
                            <td>
                                <asp:TextBox ID="tbAttnOf" runat="server" Width="400px" MaxLength="50"></asp:TextBox></td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trTitle" runat="server" visible="false">
                            <td>
                                Title</td>
                            <td>
                                <asp:TextBox ID="tbTitle" runat="server" Width="100px" MaxLength="50"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trAddr2" runat="server" visible="true">
                            <td>
                                First &amp; Last Name <span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbName" runat="server" Width="400px" MaxLength="50"></asp:TextBox></td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="tbName"
                                    ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr id="trAddr2a" runat="server" visible="true">
                            <td>
                                Company <span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbCompany" runat="server" Width="400px" MaxLength="50"></asp:TextBox></td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="rfvCompany" runat="server" ControlToValidate="tbCompany"
                                    ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr id="trJobTitle" runat="server" visible="false">
                            <td>
                                Job Title
                            </td>
                            <td>
                                <asp:TextBox ID="tbJobTitle" runat="server" Width="400px" MaxLength="50"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trAddr3" runat="server" visible="true">
                            <td>
                                Addr 1<span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbAddr1" runat="server" Width="400px" MaxLength="50"/>
                            </td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="tbAddr1" ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr id="trAddr4" runat="server" visible="true">
                            <td>
                                Addr 2</td>
                            <td>
                                <asp:TextBox ID="tbAddr2" runat="server" Width="400px" MaxLength="50"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trAddr6" runat="server" visible="true">
                            <td>
                                Town <span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbTown" runat="server" Width="400px" MaxLength="50"/>
                            </td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="tbTown" ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr id="trAddr7" runat="server" visible="true">
                            <td>
                                County / State
                            </td>
                            <td>
                                <asp:TextBox ID="tbCounty" runat="server" Width="400px" MaxLength="50"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trCtcTelNo" runat="server" visible="false">
                            <td>
                                Tel <span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbCtcTelNo" runat="server" MaxLength="50" Width="400px"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trCtcEmailAddr" runat="server" visible="false">
                            <td>
                                Email <span style="color: red">*</span></td>
                            <td>
                                <asp:TextBox ID="tbCtcEmailAddr" runat="server" MaxLength="80" Width="400px"/>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trAddr9" runat="server" visible="true">
                            <td>
                                Country <span style="color: red">*</span></td>
                            <td>
                                <asp:DropDownList ID="ddlCountry" runat="server" DataSourceID="SqlDataSourceCountries"
                                    DataTextField="CountryName" DataValueField="CountryKey" Font-Names="Verdana" Font-Size="X-Small">
                                </asp:DropDownList>&nbsp;
                                <asp:LinkButton ID="lnkbtnCountryUK" runat="server" OnClick="lnkbtnCountryUK_Click">UK</asp:LinkButton><asp:SqlDataSource
                                    ID="SqlDataSourceCountries" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
                                    SelectCommand="spASPNET_Country_GetCountries" SelectCommandType="StoredProcedure">
                                </asp:SqlDataSource>
                            </td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="ddlCountry"
                                    ErrorMessage="please select!" InitialValue="0" ValidationGroup="Address"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr runat="server" id="trCostCentre" visible="true">
                            <td>
                                Cost Centre <span style="color: red">*</span>
                            </td>
                            <td>
                                    <asp:DropDownList ID="ddlCostCentre" runat="server" Font-Names="Verdana" Font-Size="X-Small">
                                        <asp:ListItem>- please select -</asp:ListItem>
                                        <asp:ListItem>Central Marketing - 203040+64</asp:ListItem>
                                        <asp:ListItem>France - 203050+72</asp:ListItem>
                                        <asp:ListItem>Germany - 203050+66</asp:ListItem>
                                        <asp:ListItem>Germany - 200305</asp:ListItem>
                                        <asp:ListItem>GPC - 213020</asp:ListItem>
                                        <asp:ListItem>Hong Kong - 204050-85/511</asp:ListItem>
                                        <asp:ListItem>Isle of Man - 203050+70</asp:ListItem>
                                        <asp:ListItem>Italy - 203050+69</asp:ListItem>
                                        <asp:ListItem>Jersey - 203110</asp:ListItem>
                                        <asp:ListItem>Korea - TBA</asp:ListItem>
                                        <asp:ListItem>MEA Region - 204030</asp:ListItem>
                                        <asp:ListItem>Netherlands - 203050+77</asp:ListItem>
                                        <asp:ListItem>Nordic Region - 204021</asp:ListItem>
                                        <asp:ListItem>Scotland - 200305</asp:ListItem>
                                        <asp:ListItem>Singapore - TBA</asp:ListItem>
                                        <asp:ListItem>Spain 203050+73</asp:ListItem>
                                        <asp:ListItem>Switzerland - 203050+71</asp:ListItem>
                                        <asp:ListItem>Taiwan - 203050+86</asp:ListItem>
                                        <asp:ListItem>US - TBA</asp:ListItem>
                                        <asp:ListItem>UK - 203050+64</asp:ListItem>
                                        <asp:ListItem>Management - 250200</asp:ListItem>
                                        <asp:ListItem>Institutional Marketing - 204050</asp:ListItem>
                                        <asp:ListItem>Investor Services (UK) - 608530</asp:ListItem>
                                        <asp:ListItem>Investor Services (MLIIF) - 608540</asp:ListItem>
                                        <asp:ListItem>PR - 505000</asp:ListItem>
                                        <asp:ListItem>Investment Trusts - 203150</asp:ListItem>
                                        <asp:ListItem>Regulatory - code not yet defined</asp:ListItem>
                                    </asp:DropDownList>
                            </td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="rfvCostCentre" runat="server" ControlToValidate="ddlCostCentre"
                                    ErrorMessage="required!" InitialValue="- please select -" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr id="trTypeOfOrganisation" runat="server" visible="false">
                            <td>
                                Type of Organisation <span style="color: red">*</span></td>
                            <td>
                                <asp:DropDownList ID="ddlTypeOfOrganisation" runat="server" Font-Names="Verdana"
                                    Font-Size="X-Small">
                                    <asp:ListItem>- please select -</asp:ListItem>
                                    <asp:ListItem>School: Primary</asp:ListItem>
                                    <asp:ListItem>School: Middle</asp:ListItem>
                                    <asp:ListItem>School: Secondary</asp:ListItem>
                                    <asp:ListItem>Further Education / Training College</asp:ListItem>
                                    <asp:ListItem>University</asp:ListItem>
                                    <asp:ListItem>Youth service</asp:ListItem>
                                    <asp:ListItem>Drug/Alcohol Team</asp:ListItem>
                                    <asp:ListItem>Occupational health</asp:ListItem>
                                    <asp:ListItem>Health / Medical</asp:ListItem>
                                    <asp:ListItem>Armed forces</asp:ListItem>
                                    <asp:ListItem>Prison/probation service</asp:ListItem>
                                    <asp:ListItem>Police</asp:ListItem>
                                    <asp:ListItem>Licensed premises</asp:ListItem>
                                    <asp:ListItem>Local authority</asp:ListItem>
                                    <asp:ListItem>Member of the public</asp:ListItem>
                                    <asp:ListItem>Other</asp:ListItem>
                                </asp:DropDownList></td>
                            <td align="left">
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="ddlTypeOfOrganisation"
                                    ErrorMessage="please select!" InitialValue="- please select -" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr id="trAddr10" runat="server">
                            <td>
                                Special Instructions /<br />
                                Comments (optional)</td>
                            <td>
                                <asp:TextBox ID="tbSpclInstructions" runat="server" TextMode="MultiLine" Width="400px" MaxLength="1000"></asp:TextBox></td>
                            <td align="left">
                            </td>
                        </tr>
                        <tr id="trSpare" runat="server" visible="true">
                            <td>
                            </td>
                            <td>
                            </td>
                            <td align="left">
                            </td>
                        </tr>
                  <tr id="trDrinkAwareOptIn" runat="server" visible="false">
                      <td colspan="2">
                          <span style="font-size: 10pt; font-style: italic; font-family: Trebuchet MS">If you
                              would like to be kept informed about new publications and events from The Drinkaware
                              Trust, please tick this box.
                              <asp:CheckBox ID="cbOptIn" runat="server" />&nbsp;
                              We will not pass your details onto any third party.</span></td>
                      <td align="left">
                      </td>
                  </tr>
                    </table>
                    <br />
                    <table id="tblRegister" runat="server" width="90%" visible="false">
                        <tr>
                            <td style="width: 20%">
                            </td>
                            <td style="width: 65%">
                                &nbsp;
                            </td>
                            <td style="width: 15%">
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="3">
                                <asp:Label ID="lblStockAvailabilityNotificationMessage1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
                                    Text="ENTER YOUR EMAIL ADDRESS TO RECEIVE STOCK AVAILABILITY NOTIFICATIONS"/>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Email Addr:"/>
                            </td>
                            <td colspan="2">
                                <asp:TextBox ID="tbNotificationEmailAddr1" runat="server" Width="300px" MaxLength="100"/>
                                <asp:RegularExpressionValidator ID="revNotificationEmailAddr1" runat="server" ControlToValidate="tbNotificationEmailAddr1" ErrorMessage="invalid address!" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" ValidationGroup="Address" Font-Names="Verdana" Font-Size="XX-Small"/>
                                <asp:RequiredFieldValidator ID="rfvNotificationEmailAddr1" runat="server" ControlToValidate="tbNotificationEmailAddr1" ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="Address"/>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Confirm Email Addr:"/>
                            </td>
                            <td colspan="2">
                                <asp:TextBox ID="tbConfirmNotificationEmailAddr1" runat="server" Width="300px" MaxLength="100"/>
                                <asp:RegularExpressionValidator ID="revConfirmNotificationEmailAddr1" runat="server" ControlToValidate="tbConfirmNotificationEmailAddr1" ErrorMessage="invalid address!" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" ValidationGroup="Address" Font-Names="Verdana" Font-Size="XX-Small"/>
                                <asp:RequiredFieldValidator ID="rfvConfirmNotificationEmailAddr1" runat="server" ControlToValidate="tbConfirmNotificationEmailAddr1" ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="Address"/>
                                <asp:CompareValidator ID="cvNotificationEmailAddr1" runat="server" ControlToCompare="tbNotificationEmailAddr1" ControlToValidate="tbConfirmNotificationEmailAddr1" ErrorMessage="addresses must match!" ValidationGroup="Address" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:CheckBox ID="cbRegister" runat="server" AutoPostBack="True" Checked="True" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnCheckedChanged="cbRegister1_CheckedChanged" Text="Yes, send me an email when the out of stock items become available" />
                            </td>
                            <td>
                            </td>
                        </tr>
                    </table>
                    <table width="90%">
                        <tr>
                            <td style="width: 100px">
                                <asp:Label ID="lblAddressValidation" runat="server" Text="* Required" ForeColor="Red"/>
                            </td>
                            <td align="right">
                                <asp:LinkButton ID="lnkbtnBackToCategoriesFromAddressPanel" runat="server" OnClick="lnkbtnProductsByCategory_Click" SkinID="button">&nbsp;back&nbsp;to&nbsp;categories&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnBackToProductsFromAddressPanel" runat="server" OnClick="lnkbtnBackToProducts_Click" SkinID="button">&nbsp;back&nbsp;to&nbsp;products&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnAddressBook" runat="server" OnClick="btnAddressBook_Click" SkinID="button" Visible="False">&nbsp;address&nbsp;book&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnSubmitOrder" runat="server" OnClick="btnSubmitOrder_Click" SkinID="button">&nbsp;submit&nbsp;order&nbsp;</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlRegisterOnly" runat="server" Width="100%" DefaultButton="lnkbtnSubmitRequest">
                    <table width="90%">
                        <tr>
                            <td style="width: 20%">
                            </td>
                            <td style="width: 65%">
                                </td>
                            <td style="width: 15%">
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="3">
                                <asp:Label ID="lblStockAvailabilityNotificationMessage2" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
                                    Text="ENTER YOUR EMAIL ADDRESS TO RECEIVE STOCK AVAILABILITY NOTIFICATIONS"></asp:Label></td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="Label5a" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Email Addr:"></asp:Label></td>
                            <td colspan="2">
                                <asp:TextBox ID="tbNotificationEmailAddr2" runat="server" Width="300px" MaxLength="100"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="revNotificationEmailAddr2" runat="server" ControlToValidate="tbNotificationEmailAddr1"
                                    ErrorMessage="invalid address!" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" ValidationGroup="NotificationEmailAddr2" Font-Names="Verdana" Font-Size="XX-Small"></asp:RegularExpressionValidator>
                                <asp:RequiredFieldValidator ID="rfvNotificationEmailAddr2" runat="server" ControlToValidate="tbNotificationEmailAddr2"
                                    ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="NotificationEmailAddr2"></asp:RequiredFieldValidator></td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="Label6a" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Confirm Email Addr:"></asp:Label></td>
                            <td colspan="2">
                                <asp:TextBox ID="tbConfirmNotificationEmailAddr2" runat="server" Width="300px" MaxLength="100"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="revConfirmationNotificationEmailAddr2" runat="server"
                                    ControlToValidate="tbConfirmNotificationEmailAddr2" ErrorMessage="invalid address!" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*" ValidationGroup="NotificationEmailAddr2" Font-Names="Verdana" Font-Size="XX-Small"></asp:RegularExpressionValidator>
                                <asp:RequiredFieldValidator ID="rfvConfirmNotificationEmailAddr2" runat="server" ControlToValidate="tbConfirmNotificationEmailAddr2"
                                    ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="NotificationEmailAddr2"></asp:RequiredFieldValidator>
                                <asp:CompareValidator ID="cvNotificationEmailAddr2" runat="server" ControlToCompare="tbNotificationEmailAddr2"
                                    ControlToValidate="tbConfirmNotificationEmailAddr2" ErrorMessage="addresses must match!"
                                    ValidationGroup="NotificationEmailAddr2" Font-Names="Verdana" Font-Size="XX-Small"></asp:CompareValidator></td>
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
                    <table width="90%">
                        <tr>
                            <td style="width: 100px">
                            </td>
                            <td align="right">
                                <asp:LinkButton ID="lnkbtnBackToCategoriesFromAddressPanel2" runat="server" OnClick="lnkbtnProductsByCategory_Click"
                                    SkinID="button">&nbsp;back&nbsp;to&nbsp;categories&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnBackToProductsFromAddressPanel2" runat="server" OnClick="lnkbtnBackToProducts_Click"
                                    SkinID="button">&nbsp;back&nbsp;to&nbsp;products&nbsp;</asp:LinkButton>
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnSubmitRequest" runat="server" SkinID="button" OnClick="lnkbtnSubmitRequest_Click">register</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlAddressList" runat="server" DefaultButton="lnkbtnSearchAddressList"
                    Width="100%">
                    <br />
                    <asp:SqlDataSource ID="SqlDataSourceAddressList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
                        ProviderName="System.Data.SqlClient" SelectCommand="spASPNET_Address_GetGlobalAddresses"
                        SelectCommandType="StoredProcedure">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="tbSearchAddressList" DefaultValue="_" Name="SearchCriteria"
                                PropertyName="Text" Type="String" />
                            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <br />
                    <table style="width: 90%">
                        <tr>
                            <td style="width: 5px">
                            </td>
                            <td style="width: 350px">
                                Search:
                                <asp:TextBox ID="tbSearchAddressList" runat="server" Text=""></asp:TextBox>&nbsp;<asp:LinkButton
                                    ID="lnkbtnSearchAddressList" runat="server" SkinID="button" OnClick="lnkbtnSearchAddressList_Click">go</asp:LinkButton></td>
                        </tr>
                    </table>
                    <br />
                    <asp:GridView ID="gvAddressList" runat="server" DataSourceID="SqlDataSourceAddressList"
                        AllowPaging="True" AllowSorting="True" PageSize="8" GridLines="None" PagerSettings-Mode="Numeric"
                        AutoGenerateColumns="False" PagerStyle-VerticalAlign="NotSet" PagerStyle-HorizontalAlign="Center"
                        CellPadding="2">
                        <Columns>
                            <asp:TemplateField HeaderText="Name" SortExpression="Company">
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnSelectedAddress" runat="server" Text='<%# Bind("Company") %>'
                                        OnClick="lnkbtnSelectedAddress_Click">here</asp:LinkButton>
                                    <asp:HiddenField ID="hidCountryCode" runat="server" Value='<%# Bind("CountryKey") %>' />
                                    <asp:HiddenField ID="hidDefaultSpecialInstructions" runat="server" Value='<%# Bind("DefaultSpecialInstructions") %>' />
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Addr 1" SortExpression="Addr1">
                                <ItemTemplate>
                                    <asp:Label ID="lblAddr1" runat="server" Text='<%# Bind("Addr1") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Addr 2" SortExpression="Addr2">
                                <ItemTemplate>
                                    <asp:Label ID="lblAddr2" runat="server" Text='<%# Bind("Addr2") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Addr 3" SortExpression="Addr3">
                                <ItemTemplate>
                                    <asp:Label ID="lblAddr3" runat="server" Text='<%# Bind("Addr3") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Town" SortExpression="Town">
                                <ItemTemplate>
                                    <asp:Label ID="lblTown" runat="server" Text='<%# Bind("Town") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="County/State" SortExpression="State" Visible="False">
                                <ItemTemplate>
                                    <asp:Label ID="lblState" runat="server" Text='<%# Bind("State") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Post Code" SortExpression="PostCode">
                                <ItemTemplate>
                                    <asp:Label ID="lblPostCode" runat="server" Text='<%# Bind("PostCode") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Country" SortExpression="CountryName">
                                <ItemTemplate>
                                    <asp:Label ID="lblCountryName" runat="server" Text='<%# Bind("CountryName") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="AttnOf" SortExpression="AttnOf">
                                <ItemTemplate>
                                    <asp:Label ID="lblAttnOf" runat="server" Text='<%# Bind("AttnOf") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Telephone" SortExpression="Telephone">
                                <ItemTemplate>
                                    <asp:Label ID="lblTelephone" runat="server" Text='<%# Bind("Telephone") %>'></asp:Label>
                                </ItemTemplate>
                                <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                            </asp:TemplateField>
                        </Columns>
                        <PagerSettings Mode="NumericFirstLast" />
                        <PagerStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        <EmptyDataTemplate>
                            No addresses found
                        </EmptyDataTemplate>
                    </asp:GridView>
                    <br />
                    <table width="90%">
                        <tr>
                            <td style="width: 100px">
                            </td>
                            <td align="right">
                                <asp:LinkButton ID="lnkbtnBackToDeliveryAddress" runat="server" OnClick="lnkbtnBackToDeliveryAddress_Click"
                                    SkinID="button">&nbsp;back&nbsp;to&nbsp;delivery&nbsp;address&nbsp;</asp:LinkButton>
                                &nbsp;
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlSearch" runat="server" DefaultButton="btnGoSearch" Width="100%">
                    <table width="70%">
                        <tr>
                            <td>
                                <br />
                                Search for:</td>
                            <td>
                                <br />
                                <asp:TextBox ID="tbSearch" runat="server" MaxLength="30" />
                                &nbsp;&nbsp;<asp:Button ID="btnGoSearch" runat="server" Text="go" OnClick="btnGoSearch_Click" /></td>
                            <td>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlTrackAndTrace" runat="server" DefaultButton="btnCheckConsignment"
                    Width="100%">
                    <table width="80%">
                        <tr>
                            <td>
                                <br />
                                Consignment No:
                            </td>
                            <td>
                                <br />
                                <asp:TextBox ID="tbConsignmentNo" runat="server" EnableViewState="False" MaxLength="9"></asp:TextBox>
                                &nbsp;&nbsp;
                                <asp:Button ID="btnCheckConsignment" runat="server" Text="go" OnClick="btnCheckConsignment_Click" /></td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <asp:RequiredFieldValidator ID="rfvConsignmentNo" runat="server" ErrorMessage="required!"
                                    ControlToValidate="tbConsignmentNo" EnableClientScript="False" ValidationGroup="tracking"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="revConsignmentNo" runat="server" ErrorMessage="must be numeric!"
                                    ControlToValidate="tbConsignmentNo" ValidationExpression="^(\d)*$" EnableClientScript="False"
                                    ValidationGroup="tracking"></asp:RegularExpressionValidator></td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlRequestNotificationConfirmation" runat="server" Width="100%">
                    <table style="width: 100%">
                        <tr>
                            <td style="width: 100%" align="center">
                                <br />
                                <br />
                                <br />
                                <asp:Label ID="lblRequestNotificationConfirmation" runat="server" Font-Bold="True"/>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="pnlTrackingResult" runat="server" Width="100%">
                    <%=GetTracking() %>
                    <asp:Label ID="lblTrackingMessage" runat="server"></asp:Label>
                </asp:Panel>
                <asp:Panel ID="pnlBookingConfirmation" runat="server" Width="100%">
                    <p>
                        Thank you, your order is now being processed.
                    </p>
                    <p>
                        Please record this consignment number: <span class="CS_RedText">
                            <asp:Label ID="lblConsignmentKey" runat="server" Font-Bold="True"/></span>.
                    </p>
                    <p>
                        You can track this consignment using the 'Track and Trace' menu on the left of this
                        page.
                    </p>
                    <p>
                        For further assistance please e-mail Customer Services (<a href="mailto:customer_services@sprintexpress.co.uk">customer_services@sprintexpress.co.uk</a>)</p>
                </asp:Panel>
                <asp:Panel ID="pnlHelp" runat="server" Width="100%">
                </asp:Panel>
            </td>
        </tr>
    </table>
    <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
    <script type="text/javascript">
           function SB_ShowImage(value){
                window.open("show_image.aspx?Image=" + value,"ProductImage","top=10,left=10,width=610,height=610,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=yes");
           }
    </script>
</asp:Content>
