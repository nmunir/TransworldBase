<%@ Page Language="VB" MasterPageFile="~/WebForm.master" Title="Online Ordering" Inherits="PartialClassWebForm" StyleSheetTheme="Basic" %>
<%@ import Namespace=" System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Drawing.Image" %>
<%@ import Namespace="System.Drawing.Color" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Text" %>
<script runat="server">
    'from SqlDataSourceAddressList <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
    ' hack with... <asp:Parameter DefaultValue="16" Name="CustomerKey" Type="Int32" />

    '
    ' TO DO post UNIPUB release
    
    ' put base connection values in base web.config
    
    ' add keyboard navigation
    ' add client side button enabling / disabling depending on result of validation
    ' put CustomerKey in VIEWSTATE & use from there
    ' sort out duplicate stored procedures: ProductGetFromKey & GetProductFromKey (different param set)
    ' review product detail functionality esp. display of larger image
    ' allow for later addition of PDF downloads
    ' use of prod_image_folder, prod_thumb_folder
    ' integrate with web page editor
    ' skin all controls & make appearance changes easier
    ' disable Back to Product button on Basket panel if nothing in Products
    ' look at Colour.FromArgb warning & other VS warnings
    ' do something better than Server.Transfer("error.aspx") on db access error
    ' review all error handling
    ' rationalise sprocs

    ' SPROCS
    'spASPNET_WebForm_GetCategories
    'spASPNET_Product_GetSubCategories
    'spASPNET_Product_GetFromKey
    'spASPNET_Webform_GetProductFromKey
    'spASPNET_WebForm_GetTracking1
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

    Private Basket As DataTable = New DataTable()
    Private BasketView As DataView
    Dim sWebFormHomePageText As String = System.Configuration.ConfigurationManager.AppSettings("WebFormHomePageText")
    Dim sWebFormHelpPageText As String = System.Configuration.ConfigurationManager.AppSettings("WebFormHelpPageText")

    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
   
        If Not IsPostBack Then
            Dim configPath As String = "~"
            Dim config As System.Configuration.Configuration = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(configPath)
            Dim configSection As System.Web.Configuration.CompilationSection = _
              CType(config.GetSection("system.web/compilation"), _
                System.Web.Configuration.CompilationSection)
            Dim bDebug As Boolean = configSection.Debug
            If bDebug Then
                Call CheckSprocsExist()
            End If
            Call InitSessionFromConfig()

            Call GetCategories()
            Call ShowHome()

            lblBasketCount.Text = "0"
            Call AdjustBasketCountPlurality()
            bCategoryProductsFound = False
            Session.Timeout = 180
        End If
        If Not IsNumeric(Session("GenericUserKey")) Then
            Server.Transfer("timeout.aspx")
        End If
    End Sub
    
    Sub CheckSprocsExist()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("select name from sysobjects where name like 'spaspnet_webform%'", oConn)
        'oAdapter.SelectCommand.CommandText = "select name from sysobjects where name like 'spaspnet_webform%'"
        Dim dt As New DataTable

        Dim sSprocNames As New ArrayList
        sSprocNames.Add("spASPNET_WebForm_GetCategories")
        sSprocNames.Add("spASPNET_Product_GetSubCategories")
        sSprocNames.Add("spASPNET_Product_GetFromKey")          ' used in AddItemToBasket, only uses ProductKey
        sSprocNames.Add("spASPNET_Webform_GetProductFromKey")   ' used in GetProductDetail, takes CustomerKey to do EXECUTE spWAddLogisticWebHit
        sSprocNames.Add("spASPNET_WebForm_GetTracking1")
        sSprocNames.Add("spASPNET_WebForm_GetTracking2")
        sSprocNames.Add("spASPNET_WebForm_GetTracking3")
        sSprocNames.Add("spASPNET_Webform_AddBooking")
        sSprocNames.Add("spASPNET_LogisticMovement_Add")
        sSprocNames.Add("spASPNET_LogisticBooking_Complete")

        sSprocNames.Add("spASPNET_Webform_GetProductsUsingCategories")
        sSprocNames.Add("spASPNET_Webform_GetProductsUsingSearchCriteria")
        sSprocNames.Add("spASPNET_Country_GetCountries")
        Dim s As String
        For Each s In sSprocNames
            Try
                'sSQL = "select name from sysobjects where name like '%" & s & "%'"
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
    
    Sub InitSessionFromConfig()
        lCustomerKey = System.Configuration.ConfigurationManager.AppSettings("CustomerKey")
        lGenericUserKey = System.Configuration.ConfigurationManager.AppSettings ("GenericUserKey")
        Session("CustomerKey") = lCustomerKey           ' only needed for SqlDataSource - need to find out how to use Property
        Session("GenericUserKey") = lGenericUserKey     ' only needed for SqlDataSource - need to find out how to use Property
    End Sub
   
    ' C H O R E O G R A P H Y
   
    Function OkayToChangePanels()
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
   
    Sub HideAllPanels() ' ALL panel occlusion MUST come via here to ensure partially completed panel contents (eg Delivery Address) are saved
        pnlHome.Visible = False
        pnlCategorySelection.Visible = False
        pnlProductList.Visible = False
        pnlProductDetail.Visible = False
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
    End Sub
   
    Sub ShowHome()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlHome.Visible = True
            lblHeading.Text = "Home"
            lblSubHeading.Text = sWebFormHomePageText
            lblBreadcrumbLocation.Text = "home"
        End If
    End Sub

    Sub ShowCategories()
        If OkayToChangePanels() Then
           
        End If
        Call HideAllPanels()
        pnlCategorySelection.Visible = True
        lblHeading.Text = "Product Categories"
        lblSubHeading.Text = "Choose a product category, then a sub-category."
        lblBreadcrumbLocation.Text = "products by category"
    End Sub

    Sub ShowProductList()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlProductList.Visible = True
            gvCategoryProductList.DataBind()
            Dim nProductCount As String = gvCategoryProductList.Rows.Count
            Dim sProductCountPluralisation As String = ""
            If nProductCount <> 1 Then
                sProductCountPluralisation = "s"
            End If
            If nProductCount > 0 Then
                bCategoryProductsFound = True
            End If
            lblHeading.Text = "Products"
            Dim sbSubHeadingText As StringBuilder = New StringBuilder()
            sbSubHeadingText.Append(nProductCount.ToString)
            sbSubHeadingText.Append(" product")
            sbSubHeadingText.Append(sProductCountPluralisation)
            lblSubHeading.Text = sbSubHeadingText.ToString
            sbSubHeadingText.Append(" found in category ")
            sbSubHeadingText.Append("<b>")
            sbSubHeadingText.Append(sCategory)
            sbSubHeadingText.Append("</b>")
            sbSubHeadingText.Append(", subcategory ")
            sbSubHeadingText.Append("<b>")
            sbSubHeadingText.Append(sSubCategory)
            sbSubHeadingText.Append("</b>")
            sbSubHeadingText.Append(". Click the Add to Basket check box for each product you want, then click the Add to Basket button.  To view more details of the product click the View Details button.")
            lblSubHeading.Text = sbSubHeadingText.ToString
            lblBreadcrumbLocation.Text = "products by category >>> products"
        End If
    End Sub
   
    Sub ShowSearch()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlSearch.Visible = True
            lblHeading.Text = "Search for a Product"
            lblSubHeading.Text = "Search the text of product descriptions."
            lblBreadcrumbLocation.Text = "product search"
            tbSearch.Focus()
        End If
    End Sub

    Sub ShowSearchProductList()
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
            End If
        End If
    End Sub
   
    Sub ShowProductDetail()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlProductDetail.Visible = True
            lblHeading.Text = "Product Detail"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "products >>> product detail"
        End If
    End Sub
   
    Sub ShowCurrentBasket()
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
            If bCategoryProductsFound = False Then
                lnkbtnBackToProducts.Visible = False
            Else
                lnkbtnBackToProducts.Visible = True
            End If
        End If
    End Sub
   
    Sub ShowEmptyBasket()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlEmptyBasket.Visible = True
            lblHeading.Text = "Your Basket"
            lblSubHeading.Text = "Your basket is empty. Click on <b>products by category</b> to browse available items. Click on <b>product search</b> to search product descriptions."
            lblBreadcrumbLocation.Text = "basket (empty)"
        End If
    End Sub

    Sub AdjustBasketCountPlurality()
        If CLng(lblBasketCount.Text) = 1 Then
            lblBasketCountPlural.Visible = False
        Else
            lblBasketCountPlural.Visible = True
        End If
    End Sub
   
    Sub ShowAddressPanel()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlAddress.Visible = True
            lblHeading.Text = "Delivery Details"
            'lblSubHeading.Text = "Enter your name, the address to which your order is to be despatched, and any special instructions."
            lblSubHeading.Text = "Enter the post code to which your order is to be despatched, then click Find Address and select the address you require.  Enter the recipient name, any remaining address information, and any special instructions."
            lblBreadcrumbLocation.Text = "basket >>> delivery details"
            tbPostCode.Focus()
        End If
    End Sub

    Sub ShowAddressListPanel()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlAddressList.Visible = True
            tbSearchAddressList.Text = ""
            lblHeading.Text = "Address List"
            Call UpdateAddressListPanelLegends()
        End If
    End Sub
    
    Sub UpdateAddressListPanelLegends()
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

    Sub ShowTrackAndTrace()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlTrackAndTrace.Visible = True
            lblHeading.Text = "Track And Trace"
            lblSubHeading.Text = "Enter the consignment number displayed when your order was placed."
            lblBreadcrumbLocation.Text = "track and trace"
            tbConsignmentNo.Focus()
        End If
    End Sub
   
    Sub ShowTrackingResult()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlTrackingResult.Visible = True
            lblHeading.Text = "Tracking Result"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "track and trace >>> tracking result"
        End If
    End Sub
   
    Sub ShowBookingConfirmation()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlBookingConfirmation.Visible = True
            lblHeading.Text = "Booking Confirmation"
            lblSubHeading.Text = ""
            lblBreadcrumbLocation.Text = "basket >>> delivery details >>> booking confirmation"
        End If
    End Sub
   
    Sub ShowHelp()
        If OkayToChangePanels() Then
            Call HideAllPanels()
            pnlHelp.Visible = True
            lblHeading.Text = "Help"
            lblSubHeading.Text = sWebFormHelpPageText
            lblBreadcrumbLocation.Text = "help"
        End If
    End Sub
   
    ' C A T E G O R Y   &   S U B - C A T E G O R Y   L I S T S

    Sub btn_ReturnToMainPanel_click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        ShowProductList()
    End Sub
   
    Sub lnkbtn_ShowProductsByCategory_click(ByVal sender As Object, ByVal e As CommandEventArgs)   ' user clicked on a sub-category so show list of products
        sSubCategory = CStr(e.CommandArgument)  ' Put into VIEWSTATE
        ShowProductList()
    End Sub

    Sub rptrCategories_Item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim Colour As System.Drawing.Color
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = Colour.FromArgb(131, 148, 140)
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = Colour.FromArgb(242, 173, 13) 'selected
        lblSubCategoryHeading.Visible = True
    End Sub

    Sub lnkbtn_ShowSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        sCategory = CStr(e.CommandArgument)
        rptrSubCategories.Visible = True
        GetSubCategories()
    End Sub

    Sub btn_ShowCategories_click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        Call DisplayCategories()
    End Sub

    Sub DisplayCategories()
        Dim Colour As System.Drawing.Color
        Dim item As RepeaterItem
        For Each item In rptrCategories.Items
            Dim x As LinkButton = CType(item.Controls(3), LinkButton)
            x.ForeColor = Colour.FromArgb(131, 148, 140)
        Next
        rptrSubCategories.Visible = False
        lblSubCategoryHeading.Visible = False
        Call ShowCategories()
    End Sub

    Sub GetCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_WebForm_GetCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = lCustomerKey
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

    Sub GetSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        ' Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubCategories", oConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_WebForm_GetSubCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = sCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = lCustomerKey   ' Session("CustomerKey")

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

    ' B A S K E T

    Sub AddItemToBasket(ByVal sProductKey As String)
        Dim dr As DataRow
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetFromKey", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int)
        oParamProductKey.Value = CLng(sProductKey)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
           
            If IsNothing(Session("BasketData")) Then            ' create a new Basket structure & assign it to session variable
                Basket = New DataTable()
                Basket.Columns.Add(New DataColumn("ProductKey", GetType(String)))
                Basket.Columns.Add(New DataColumn("ProductCode", GetType(String)))
                Basket.Columns.Add(New DataColumn("ProductDate", GetType(String)))
                Basket.Columns.Add(New DataColumn("Description", GetType(String)))
                Basket.Columns.Add(New DataColumn("BoxQty", GetType(String)))
                Basket.Columns.Add(New DataColumn("UnitWeightGrams", GetType(Double)))
                Basket.Columns.Add(New DataColumn("UnitValue", GetType(Double)))
                Basket.Columns.Add(New DataColumn("QtyAvailable", GetType(Long)))
                Basket.Columns.Add(New DataColumn("QtyToPick", GetType(Long)))
                Basket.Columns.Add(New DataColumn("QtyRequested", GetType(Long)))
                Basket.Columns.Add(New DataColumn("PDFFileName", GetType(String)))
                Basket.Columns.Add(New DataColumn("OriginalImage", GetType(String)))
                Basket.Columns.Add(New DataColumn("ThumbNailImage", GetType(String)))
                Session("BasketData") = Basket
            End If

            Basket = Session("BasketData")                              ' init Basket from session variable
            BasketView = New DataView(Basket)                           ' create a DataView of it
            BasketView.RowFilter = "ProductKey='" & sProductKey & "'"   ' is selected product already in Basket?
            If BasketView.Count = 0 Then                                ' no, so add it to Basket
                dr = Basket.NewRow()
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
                Basket.Rows.Add(dr)
                lblBasketCount.Text = CLng(lblBasketCount.Text) + 1             ' increment Basket item count
                Call AdjustBasketCountPlurality()
                Session("BasketData") = Basket                                  ' store Basket in session variable
                BasketView.RowFilter = ""
            End If
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Sub
   
    Sub RemoveItemFromBasket(ByVal sProductKey As String)
        Basket = Session("BasketData")                                          ' get Basket from session variable
        BasketView = New DataView(Basket)                                       ' create a DataView of Basket
        BasketView.RowFilter = "ProductKey='" & sProductKey & "'"               ' set filter to selected record
        If BasketView.Count > 0 Then                                            ' if record is present
            BasketView.Delete(0)                                                ' remove it
            lblBasketCount.Text = CLng(lblBasketCount.Text) - 1                 ' decrement Basket item count
            Call AdjustBasketCountPlurality()
        End If
        BasketView.RowFilter = ""
        Session("BasketData") = Basket                                          ' store Basket back in session variable
    End Sub
   
    Protected Sub btnRemoveItemFromBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call btnRemoveItemFromBasket_Click()
    End Sub
   
    Protected Sub btnRemoveItemFromBasket_Click()                               ' NB: each item in Basket has a Remove checkbox and associated hidden field that holds the Product Key
        Dim cbRemoveItem As CheckBox
        Dim hidProductKey As HiddenField
        For Each row As GridViewRow In gvBasket.Rows                            ' go through the Basket remove any checked items
            cbRemoveItem = row.FindControl("cbRemoveItemFromBasket")
            If cbRemoveItem.Checked = True Then
                hidProductKey = row.FindControl("hidProductKey")
                RemoveItemFromBasket(hidProductKey.Value)
            End If
        Next
        Call BindBasketGrid("ProductCode")

    End Sub

    Sub BindBasketGrid(ByVal SortField As String)                               ' bind Basket GridView to Basket held in session variable
        If Not IsNothing(Session("BasketData")) Then
            Basket = Session("BasketData")
            BasketView = New DataView(Basket)
            BasketView.Sort = SortField
            If BasketView.Count > 0 Then
                gvBasket.DataSource = BasketView
                gvBasket.DataBind()
                ShowCurrentBasket()
            Else
                ShowEmptyBasket()
            End If
        Else
            ShowEmptyBasket()
        End If
    End Sub

    ' L H S   M E N U   H Y P E R L I N K S
   
    Protected Sub lnkbtnHome_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowHome()
    End Sub

    Protected Sub lnkbtnProductsByCategory_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub SqlDataSourceCategoryProductList_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs)
        If Len(ViewState("SubCategory")) > 0 Then
            e.Command.Parameters("@Category").Value = ViewState("Category")
            e.Command.Parameters("@SubCategory").Value = ViewState("SubCategory")
        End If
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
        Page.Validate("Address")
        If Page.IsValid Then
            If SubmitOrder() = True Then
                Call ShowBookingConfirmation()
                Call TidyUp()
            Else
                '
            End If
        Else
        End If
    End Sub

    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvSearchProductList.DataBind()
        Call ShowSearchProductList()
    End Sub

    Protected Sub SqlDataSourceSearchProductList_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs)
        '        If Len(ViewState("SubCategory")) > 0 Then
        ' e.Command.Parameters("@Category").Value = ViewState("Category")
        ' e.Command.Parameters("@SubCategory").Value = ViewState("SubCategory")
        ' End If
    End Sub

    Sub GetProductDetail(ByVal sProductKey As String)
        Dim oDataReader As SqlDataReader

        Dim sVirtualThumbFolder = System.Configuration.ConfigurationManager.AppSettings("Virtual_Thumb_URL")
        Dim sVirtualJPGFolder = System.Configuration.ConfigurationManager.AppSettings("Virtual_JPG_URL")

        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Webform_GetProductFromKey", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCustomerKey As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
        oParamCustomerKey.Value = lCustomerKey
        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int, 4)
        oParamProductKey.Value = CLng(sProductKey)
        Session("ProductKey") = CLng(sProductKey)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("ProductCode")) Then
                lblProductCode.Text = oDataReader("ProductCode")
            End If
            If IsDBNull(oDataReader("ProductDescription")) Then
                lblSubHeading.Text = ""
            Else
                lblSubHeading.Text = oDataReader("ProductDescription")
            End If
            If Not IsDBNull(oDataReader("ThumbNailImage")) Then     ' eg when  "Virtual_Thumb_URL" ="./prod_images/thumbs/" then result is ./prod_images/thumbs/13256.jpg
                imgThumbNail.ImageUrl = sVirtualThumbFolder & oDataReader("ThumbNailImage")
                If oDataReader("OriginalImage") = "blank_image.jpg" Then
                    hlnk_OriginalImage.Visible = False
                Else                                                ' original image; eg when  "Virtual_Thumb_URL" ="./prod_images/jpgs/" then result is ./prod_images/jpgs/13256.jpg
                    hlnk_OriginalImage.Visible = True
                    hlnk_OriginalImage.NavigateUrl = sVirtualJPGFolder & oDataReader("OriginalImage")
                End If
            End If
        Catch ex As SqlException
            Server.Transfer("error.aspx")
            lblError.Text = ex.ToString
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Sub
   
    Function GetPDFLink() As String
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
   

    Protected Sub btnCheckConsignment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("tracking")
        If Page.IsValid Then
            Call ShowTrackingResult()
        End If
    End Sub

    Function GetTracking() As String
        Dim sbHTML As StringBuilder = New StringBuilder()
        Dim t As String = "                                        " 'tabs for indenting
        Dim sConsignee As String = ""
        Dim sNOP As String = ""
        Dim sWeight As String = ""
        Dim sPODStatus As String = ""
        Dim lConsignmentKey As Long
        Dim oDataReader1 As SqlDataReader
        Dim bRecordFound As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_WebForm_GetTracking1", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As New SqlParameter("@ConsignmentNo", SqlDbType.NVarChar, 50)
        oCmd.Parameters.Add(oParam)
        oParam.Value = tbConsignmentNo.Text
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
            Server.Transfer("error.aspx ")
            'lblError.Text = ex.ToString
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
                'lblError.Text = ex.ToString
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
                    'dgrdConsignmentItems.Visible = False
                End If
            Catch ex As SqlException
                Server.Transfer("error.aspx")
                'lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            Call ShowTrackingResult()
        Else
            lblTrackingMessage.Text = "Consignment not found."
            'Call ShowTrackAndTrace()
        End If
        Return sbHTML.ToString()
    End Function

    Function SubmitOrder()
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim bBookingFailed As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_Webform_AddBooking", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = lGenericUserKey
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = lCustomerKey
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
        param7.Value = tbAddr1.Text
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param8.Value = tbAddr2.Text
        oCmdAddBooking.Parameters.Add(param8)
        Dim param9 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param9.Value = tbAddr3.Text
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
        param15.Value = tbCtcTelNo.Text
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param16.Value = tbSpclInstructions.Text
        oCmdAddBooking.Parameters.Add(param16)
   
        Dim param21 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
        param21.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        param22.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param22)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "AddBooking")
            'Create the Master (stock booking) record
            oCmdAddBooking.Connection = oConn
            oCmdAddBooking.Transaction = oTrans
            oCmdAddBooking.ExecuteNonQuery()
            'Output parameter contains the new Booking Key
            lBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value)
            lConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value)
            If lBookingKey > 0 Then
                'Create the child (stock movement) records
                Basket = Session("BasketData")
                BasketView = Basket.DefaultView
                Dim ProductItem As DataRowView
                If BasketView.Count > 0 Then
                    For Each ProductItem In BasketView
                        Dim lProductKey As Long = CLng(ProductItem("ProductKey"))
                        Dim lPickQuantity As Long = CLng(ProductItem("QtyRequested"))
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                        param51.Value = lGenericUserKey
                        oCmdAddStockItem.Parameters.Add(param51)
                        Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                        param52.Value = lCustomerKey
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
                    Next
                    ' added child records; mark stock booking as complete and ready for autopicker
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
        Return Not bBookingFailed
    End Function
   
    Protected Sub lnkbtnBackToProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductList()
    End Sub

    Protected Sub lnkbtnProceedToCheckout_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("Basket")
        If Page.IsValid Then
            'Call RecordQuantities()
            Call ShowAddressPanel()
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
        Basket = Session("BasketData")
        BasketView = New DataView(Basket)

        For Each gvr In gvBasket.Rows
            hidProductKey = gvr.FindControl("hidProductKey")
            hidQtyAvailable = gvr.FindControl("hidQtyAvailable")
            lblProductCode = gvr.FindControl("lblProductCode")
            tbQuantity = gvr.FindControl("tbQuantity")
            x = tbQuantity.Text
            BasketView.RowFilter = "ProductKey='" & hidProductKey.Value & "'"
            If BasketView.Count = 1 AndAlso IsNumeric(tbQuantity.Text) Then
                If CLng(hidQtyAvailable.Value) >= CLng(tbQuantity.Text) Then
                    BasketView(0).Item("QtyRequested") = CLng(tbQuantity.Text)
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
   
    Protected Sub lnkbtnViewDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hidProductKey As HiddenField = sender.NamingContainer.FindControl("hidProductKey")
        Call GetProductDetail(hidProductKey.Value)
        Call ShowProductDetail()
    End Sub
   
    Sub TidyUp()
        Session.Clear()
        InitSessionFromConfig()
        bCategoryProductsFound = False
        lblBasketCount.Text = "0"
        Call AdjustBasketCountPlurality()
        ' need to clear out VIEWSTATE too?
    End Sub

    Protected Sub gvProductList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim cbAddToOrder As CheckBox
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim nQtyAvailable As Integer = CInt(DataBinder.Eval(e.Row.DataItem, "QtyAvailable"))
            If nQtyAvailable = 0 Then
                cbAddToOrder = e.Row.FindControl("cbAddToOrder")
                cbAddToOrder.Enabled = False
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
        Dim lbl As Label, hid As HiddenField, gvr As GridViewRow
        
        lnkbtnSelAddr = CType(sender, LinkButton)
        gvr = lnkbtnSelAddr.NamingContainer
        Dim x As Object
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
        tbAddr3.Text = CType(gvr.FindControl("lblAddr3"), Label).Text
        tbTown.Text = CType(gvr.FindControl("lblTown"), Label).Text
        tbCounty.Text = CType(gvr.FindControl("lblState"), Label).Text
        tbPostCode.Text = CType(gvr.FindControl("lblPostCode"), Label).Text
        tbCtcTelNo.Text = CType(gvr.FindControl("lblTelephone"), Label).Text
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

    Protected Sub lnkbtnFindAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPostCode.Text = tbPostCode.Text.Trim.ToUpper
        trAddr1.Visible = False
        trAddr2.Visible = False
        trAddr3.Visible = False
        trAddr4.Visible = False
        trAddr5.Visible = False
        trAddr6.Visible = False
        trAddr7.Visible = False
        trAddr8.Visible = False
        trAddr9.Visible = False
        trAddr10.Visible = False
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

            'Add the new items to the list
            If Not objInterimResults.Results Is Nothing Then
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
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            Try
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
        tbAttnOf.Focus()

        If objAddressResults.IsError Then
            lblLookupError.Text = objAddressResults.ErrorMessage
        Else
            objAddress = objAddressResults.Results(0)

            tbName.Text = objAddress.OrganisationName
            tbAddr1.Text = objAddress.Line1
            tbAddr2.Text = objAddress.Line2
            tbAddr3.Text = objAddress.Line3
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
        Call ShowAllAddressFields()
    End Sub
    
    Protected Sub ShowAllAddressFields()
        trAddr1.Visible = True
        trAddr2.Visible = True
        trAddr3.Visible = True
        trAddr4.Visible = True
        trAddr5.Visible = True
        trAddr6.Visible = True
        trAddr7.Visible = True
        trAddr8.Visible = True
        trAddr9.Visible = True
        trAddr10.Visible = True
        trPostcodeLookupResults.Visible = False
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
        'If Not bContinue Then
        ' Response.Flush()
        ' Response.End()
        ' End If
    End Sub
    
</script>

<asp:Content ID="ContentBreadcrumbs" ContentPlaceHolderID="ContentPlaceHolderBreadcrumbs" Runat="Server">
    <table  border="0" cellpadding="0" cellspacing="0" style="width: 100%; height: 100%">
        <tr>
            <td nowrap="nowrap" align="left">
&nbsp;you are here:
                <asp:Label ID="lblBreadcrumbLocation" runat="server" Text="home"></asp:Label></td>
            <td nowrap="nowrap">
                &nbsp; - your basket contains <asp:Label runat="server" id="lblBasketCount">0</asp:Label>
                item<asp:Label runat="server" id="lblBasketCountPlural">s</asp:Label> -</td>
        </tr>
    </table>
</asp:Content>

<asp:Content ID="ContentNavigation" ContentPlaceHolderID="ContentPlaceHolderNavigation" runat="Server">
 <table border="0" cellspacing="0" cellpadding="0" id="navigation" width="185px">
        <tr>
          <td width="135"><a href="" class="navText">
             <asp:LinkButton ID="lnkbtnHome"
                             runat="server"
                             Width="135px"
                             OnClick="lnkbtnHome_Click">
                             &nbsp;home&nbsp;</asp:LinkButton></a></td>
        </tr>
        <tr>
          <td width="135"><a href="" class="navText">
              <asp:LinkButton ID="lnkbtnProductsByCategory"
                              runat="server"
                              Width="135px"
                              OnClick="lnkbtnProductsByCategory_click">
                              &nbsp;products&nbsp;by&nbsp;category&nbsp;</asp:LinkButton>
          </a></td>
        </tr>
        <tr>
          <td width="135"><a href="" class="navText">
              <asp:LinkButton ID="lnkbtnProductSearch"
                              runat="server"
                              Width="135px"
                              OnClick="lnkbtnProductSearch_Click">
                              &nbsp;product&nbsp;search&nbsp;</asp:LinkButton></a></td>
        </tr>
        <tr>
          <td width="135"><a href="" class="navText">
              <asp:LinkButton ID="lnkbtnMyOrder"
                              runat="server"
                              Width="135px"
                              OnClick="lnkbtnMyOrder_Click">
                              &nbsp;your&nbsp;basket&nbsp;</asp:LinkButton></a></td>
        </tr>
        <tr>
          <td width="135"><a href="" class="navText">
              <asp:LinkButton ID="lnkbtnTrackAndTrace"
                              runat="server"
                              Width="135px"
                              OnClick="lnkbtnTrackAndTrace_Click">
                              &nbsp;track&nbsp;&&nbsp;trace&nbsp;</asp:LinkButton></a></td>
        </tr>
        <tr>
          <td width="135"><a href="" class="navText">
              <asp:LinkButton ID="lnkbtnHelpWithOrdering"
                              runat="server"
                              Width="135px"
                              OnClick="lnkbtnHelpWithOrdering_Click">
                              &nbsp;help&nbsp;</asp:LinkButton></a></td>
        </tr>
      </table>
</asp:Content>

<asp:Content ID="ContentMain" ContentPlaceHolderID="ContentPlaceHolderMain" runat="Server">
 <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <tr>
          <td class="pageName" style="height: 20px">
              <asp:Label ID="lblHeading" runat="server" Text=""></asp:Label></td>
  </tr>

  <tr>
          <td class="bodyText" style="height: 15px"><br /><p>
              <asp:Label ID="lblSubHeading" runat="server" Text=""></asp:Label>&nbsp;</p>

  </td>
        </tr>
  <tr>
          <td class="bodyText">
            <asp:Panel id="pnlHome" runat="server" visible="False">
                <asp:Table id="tblHome" runat="server" Width="100%" >
                    <asp:TableRow runat="server">
                        <asp:TableCell wrap="False" HorizontalAlign="Center" runat="server">
                            <br />
                        </asp:TableCell>
                    </asp:TableRow>
                    <asp:TableRow runat="server" >
                        <asp:TableCell HorizontalAlign="Center" runat="server">
                            <br />
                            <br />
                            <br /><br />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </asp:Panel>
         <asp:Panel id="pnlCategorySelection" runat="server" visible="True" >
            <asp:Table id="tblCategorySelection" runat="server" Width="100%"  >
                <asp:TableRow>
                    <asp:TableCell Width="5%"></asp:TableCell>
                    <asp:TableCell Width="30%" Wrap="False">
                        <br />
                    </asp:TableCell>
                    <asp:TableCell Width="65%"></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                        <asp:Label ID="Label14"
                                   runat="server"
                                   ForeColor="#72837b"
                                   font-bold="True">&nbsp;&nbsp; Category</asp:Label>
                        <br/>
                        <br/>
                        <asp:Repeater runat="server"
                                      ID="rptrCategories"
                                      OnItemCommand="rptrCategories_Item_click">
                            <ItemTemplate>
                                <asp:Image ID="Image1"
                                           runat="server"
                                           ImageUrl="./images/greycircle.gif"></asp:Image>
                                <asp:LinkButton ID="lnkbtnShowSubCategories"
                                                runat="server"
                                                OnCommand="lnkbtn_ShowSubCategories_click"
                                                CommandArgument='<%# Container.DataItem("Category")%>'
                                                Text='<%# Container.DataItem("Category")%>'
                                                ForeColor="#83948C"
                                                ></asp:LinkButton>
                                <br /><br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                        <asp:Label runat="server"
                                   id="lblSubCategoryHeading"
                                   ForeColor="#72837b"
                                   Font-bold="True"
                                   Visible="False">&nbsp;&nbsp; Sub-Category</asp:Label>
                        <br/>
                        <br/>
                        <asp:Repeater runat="server" Visible="False" ID="rptrSubCategories">
                            <ItemTemplate>
                                <asp:Image ID="Image2"
                                           runat="server"
                                           ImageUrl="./images/greycircle.gif"
                                           ></asp:Image>
                                <asp:LinkButton ID="lnkbtnShowProducsByCategory"
                                                runat="server"
                                                OnCommand="lnkbtn_ShowProductsByCategory_click"
                                                CommandArgument='<%# Container.DataItem("SubCategory")%>'
                                                Text='<%# Container.DataItem("SubCategory")%>'
                                                ForeColor="#83948C"
                                                >
                                                </asp:LinkButton>
                                <br /><br />
                            </ItemTemplate>
                        </asp:Repeater>
                        <br />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
         <asp:Panel id="pnlProductList" Runat="server" Visible="False">
         <br />
          <asp:SqlDataSource ID="SqlDataSourceCategoryProductList" runat="server"
           ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
           SelectCommand="spASPNET_Webform_GetProductsUsingCategories"
           SelectCommandType="StoredProcedure"
           OnSelecting="SqlDataSourceCategoryProductList_Selecting">
              <SelectParameters>
                  <asp:SessionParameter DefaultValue="-1" Name="CustomerKey" SessionField="CustomerKey"
                      Type="Int32" />
                  <asp:Parameter Name="Category" Type="String" />
                  <asp:Parameter Name="SubCategory" Type="String" />
                  <asp:SessionParameter DefaultValue="-1" Name="GenericUserKey" SessionField="GenericUserKey"
                    Type="Int32" />
              </SelectParameters>
          </asp:SqlDataSource>
             <asp:GridView ID="gvCategoryProductList"
                  runat="server"
                  AllowPaging="True"
                  AutoGenerateColumns="False"
                  DataKeyNames="LogisticProductKey"
                  DataSourceID="SqlDataSourceCategoryProductList"
                  PageSize="5"
                  GridLines="None"
                  PagerSettings-Mode="NextPreviousFirstLast"
                  OnRowDataBound="gvProductList_RowDataBound">
                 <Columns>
                     <asp:TemplateField>
                         <ItemTemplate>
                             &nbsp;&nbsp;<asp:LinkButton ID="lnkbtnViewDetails" runat="server"
                             OnClick="lnkbtnViewDetails_Click"
                             SkinID="button">&nbsp;View&nbsp;Details&nbsp;</asp:LinkButton>&nbsp;
                         </ItemTemplate>
                         <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                     </asp:TemplateField >
                     <asp:TemplateField HeaderText="Product Code">
                         <ItemTemplate>
                             <asp:HiddenField ID="hidProductKey" Value='<%# Bind("LogisticProductKey") %>' runat="server" />
                             <asp:Label ID="Label2" runat="server" Text='<%# Bind("ProductCode") %>' Width="89px"></asp:Label>
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" Height="50px" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px"/>
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Product Description" SortExpression="ProductDescription">
                         <ItemTemplate>
                             &nbsp;<asp:Label ID="Label1" runat="server" Text='<%# Bind("ProductDescription") %>' Width="400px"></asp:Label>
                         </ItemTemplate>
                         <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Qty Available" SortExpression="QtyAvailable">
                         <ItemTemplate>
                             <asp:Label ID="Label3" runat="server" Text='<%# Bind("QtyAvailable") %>'></asp:Label>
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="&#160;Add to Basket&#160;">
                         <ItemTemplate>
                             <asp:CheckBox ID="cbAddToOrder" runat="server" Width="70px" />
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                     </asp:TemplateField>
                 </Columns>
                 <PagerSettings Mode="NumericFirstLast" />
                 <PagerStyle HorizontalAlign="Center" VerticalAlign="Middle" />
             </asp:GridView>
             <br />
            <asp:Table id="Table1" runat="server" Width="100%" >
                <asp:TableRow ID="TableRow3" runat="server" VerticalAlign="Middle" >
                    <asp:TableCell ID="TableCell7" width="10%" runat="server">&nbsp;</asp:TableCell>
                    <asp:TableCell ID="TableCell8" width="30%" runat="server">&nbsp;
                      </asp:TableCell>
                    <asp:TableCell ID="TableCell9" width="60%" runat="server" VerticalAlign="Middle" HorizontalAlign="Center" >
                         <asp:LinkButton ID="lnkbtnAddToOrder" runat="server"
                                         OnClick="lnkbtnAddToOrderFromCategories_Click"
                                         SkinID="button">&nbsp;Add&nbsp;to&nbsp;Basket&nbsp;</asp:LinkButton>
                        </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
         </asp:Panel>
        
         <asp:Panel id="pnlSearchProductList" Runat="server" Visible="False">
         <br />
          <asp:SqlDataSource ID="SqlDataSourceSearchProductList" runat="server"
           ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
           SelectCommand="spASPNET_Webform_GetProductsUsingSearchCriteria"
           SelectCommandType="StoredProcedure"
           OnSelecting="SqlDataSourceSearchProductList_Selecting"
           >
              <SelectParameters>
                  <asp:SessionParameter DefaultValue="-1" Name="CustomerKey" SessionField="CustomerKey"
                    Type="Int32" />
                  <asp:ControlParameter ControlID="tbSearch" DefaultValue="_" Name="SearchCriteria"
                      PropertyName="Text" Type="String" />
                  <asp:SessionParameter DefaultValue="-1"
                                        Name="GenericUserKey"
                                        SessionField="GenericUserKey"
                                        Type="Int32" />
              </SelectParameters>
          </asp:SqlDataSource>

             <asp:GridView ID="gvSearchProductList"
                  runat="server"
                  AllowPaging="True"
                  AutoGenerateColumns="False"
                  DataKeyNames="LogisticProductKey"
                  DataSourceID="SqlDataSourceSearchProductList"
                  PageSize="5"
                  GridLines="None"
                  PagerSettings-Mode="NextPreviousFirstLast"
                  OnRowDataBound="gvProductList_RowDataBound"
                  >
                 <Columns>
                     <asp:TemplateField  ItemStyle-BorderWidth="1px" ItemStyle-BorderStyle="Dotted" ItemStyle-BorderColor="#dddddd" >
                         <ItemTemplate>
                             &nbsp;&nbsp;<asp:LinkButton ID="lnkbtnViewDetails" runat="server"
                             OnClick="lnkbtnViewDetails_Click"
                             SkinID="button">&nbsp;View&nbsp;Details&nbsp;</asp:LinkButton>&nbsp;
                         </ItemTemplate>
                     </asp:TemplateField >
                     <asp:TemplateField HeaderText="Product Code" ItemStyle-BorderWidth="1px" ItemStyle-BorderStyle="Dotted" ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                             <asp:HiddenField ID="hidProductKey" Value='<%# Bind("LogisticProductKey") %>' runat="server" />
                             <asp:Label ID="Label2" runat="server" Text='<%# Bind("ProductCode") %>' Width="89px"></asp:Label>
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" Height="50px"/>
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Product Description" SortExpression="ProductDescription" ItemStyle-BorderWidth="1px" ItemStyle-BorderStyle="Dotted" ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                             &nbsp;<asp:Label ID="Label1" runat="server" Text='<%# Bind("ProductDescription") %>' Width="400px"></asp:Label>
                         </ItemTemplate>
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Qty Available" SortExpression="QtyAvailable" ItemStyle-BorderWidth="1px" ItemStyle-BorderStyle="Dotted" ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                             <asp:Label ID="Label3" runat="server" Text='<%# Bind("QtyAvailable") %>'></asp:Label>
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="&nbsp;Add to Basket&nbsp;" ItemStyle-BorderWidth="1px" ItemStyle-BorderStyle="Dotted" ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                             <asp:CheckBox ID="cbAddToOrder" runat="server" Width="70px" />
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                 </Columns>
                 <PagerSettings Mode="NumericFirstLast" />
                 <PagerStyle HorizontalAlign="Center" VerticalAlign="Middle" />
             </asp:GridView>
             <br />
            <asp:Table id="tabProductsFooter" runat="server" Width="100%" >
                <asp:TableRow ID="TableRow2" runat="server" VerticalAlign="Middle" >
                    <asp:TableCell ID="TableCell4" width="10%" runat="server"></asp:TableCell>
                    <asp:TableCell ID="TableCell5" width="30%" runat="server">
                      </asp:TableCell>
                    <asp:TableCell ID="TableCell6" width="60%" runat="server" VerticalAlign="Middle" HorizontalAlign="Center" >
             <asp:LinkButton ID="lnkbtnSearchAddToOrder" runat="server"
                             OnClick="lnkbtnAddToOrderFromSearch_Click"
                             SkinID="button">&nbsp;Add&nbsp;to&nbsp;Basket&nbsp;</asp:LinkButton>
                        </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            
         </asp:Panel>
             <asp:Panel id="pnlProductDetail" runat="server">
              <table width="95%">
                  <tr>
                      <td >
                      </td>
                      <td >
                      </td>
                      <td >
                      </td>
                  </tr>
              </table>
              Product Code: <asp:Label ID="lblProductCode" runat="server" Text="Label"></asp:Label>
              <br /><br />
                 <asp:Image ID="imgThumbNail" runat="server" />
              <br />
                 <asp:HyperLink ID="hlnk_OriginalImage"
                                runat="server"
                                Target="_blank"
                                >show larger image</asp:HyperLink>
              <br />
                 <%=GetPDFLink() %>
              </asp:Panel>

         <asp:Panel id="pnlBasket"
                    runat="server"
                    visible="False">
            <asp:Table id="tabBasketHeader"
                       runat="server"
                       Width="100%">
                <asp:TableRow>
                    <asp:TableCell width="10%">&nbsp;</asp:TableCell>
                    <asp:TableCell width="80%">
                    </asp:TableCell>
                    <asp:TableCell width="10%"></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
             <asp:GridView ID="gvBasket"
                           runat="server"
                           AutoGenerateColumns="False"
                           Width="100%"
                           ShowFooter="True"
                           GridLines="None">
                 <Columns>
                     <asp:TemplateField HeaderText="remove item"
                                        ItemStyle-BorderWidth="1px"
                                        ItemStyle-BorderStyle="Dotted"
                                        ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate >
                             <asp:CheckBox ID="cbRemoveItemFromBasket"
                                           runat="server" />
                             <asp:HiddenField ID="hidProductKey"
                                              Value='<%# Bind("ProductKey") %>'
                                              runat="server" />                        
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                         <FooterTemplate>
                      <img src="./images/1TxparentPixel.gif"
                           height="20px"
                           width="1px" />
                           <asp:LinkButton ID="lnkbtnRemoveItemFromBasket"
                                           runat="server"
                                           OnClick="lnkbtnRemoveItemFromBasket_Click" SkinID="button">&nbsp;remove&nbsp;</asp:LinkButton>

                         </FooterTemplate>
                         <FooterStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Product Code"
                                        ItemStyle-BorderWidth="1px"
                                        ItemStyle-BorderStyle="Dotted"
                                        ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                             &nbsp;<asp:Label ID="lblProductCode"
                                              runat="server"
                                              Text='<%# Bind("ProductCode") %>'
                                              Width="90px"></asp:Label>
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Product Description"
                                        ItemStyle-BorderWidth="1px"
                                        ItemStyle-BorderStyle="Dotted"
                                        ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                        <asp:Label ID="lblProductDescription"
                                   runat="server"
                                   Text='<%# Bind("Description") %>'
                                   Width="400px"></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                     <asp:TemplateField HeaderText="Quantity"
                                        ItemStyle-BorderWidth="1px"
                                        ItemStyle-BorderStyle="Dotted"
                                        ItemStyle-BorderColor="#dddddd">
                         <ItemTemplate>
                         <ItemStyle width=45px>
                         <asp:TextBox ID="tbQuantity"
                                      runat="server"
                                      MaxLength="4"
                                      Text='<% # Bind("QtyRequested") %>'
                                      Width="22px">
                           </asp:TextBox><asp:RequiredFieldValidator ID="RequiredFieldValidatorBasket"
                                                         runat="server"
                                                         ControlToValidate="tbQuantity"
                                                         ValidationGroup="Basket"
                                                         ErrorMessage="&nbsp;required!&nbsp;"
                                                         Display="Dynamic">
                             </asp:RequiredFieldValidator><asp:RangeValidator ID="RangeValidatorBasketQuantity"
                                                 runat="server"
                                                 ControlToValidate="tbQuantity"
                                                 ErrorMessage="invalid quantity!"
                                                 MinimumValue="1"
                                                 MaximumValue="99999"
                                                 Display="Dynamic"
                                                 ValidationGroup="Basket"></asp:RangeValidator>
                             <asp:HiddenField ID="hidQtyAvailable"
                                              Value='<%# Bind("QtyAvailable") %>'
                                              runat="server" />                        
                         </ItemTemplate>
                         <ItemStyle HorizontalAlign="Center" />
                     </asp:TemplateField>
                 </Columns>
             </asp:GridView>
            <asp:Table id="tblBasketFooter" runat="server" Width="100%" >
                <asp:TableRow ID="TableRow1" runat="server" VerticalAlign="Middle" >
                    <asp:TableCell ID="TableCell1" width="10%" runat="server"></asp:TableCell>
                    <asp:TableCell ID="TableCell2" width="30%" runat="server">
                        <asp:Label ID="lblInsufficientQuantityAvailable" runat="server" Text="Insufficient quantity available!" ForeColor="Red" Visible="false" EnableViewState="false"></asp:Label>
                      </asp:TableCell>
                    <asp:TableCell ID="TableCell3" width="60%" runat="server" VerticalAlign="Middle" >
                        <asp:LinkButton ID="lnkbtnBackToCategories" runat="server"
                        OnClick="lnkbtnProductsByCategory_Click" SkinID="button">&nbsp;Back&nbsp;to&nbsp;Categories&nbsp;</asp:LinkButton>
                        &nbsp;
                        <asp:LinkButton ID="lnkbtnBackToProducts" runat="server"
                        OnClick="lnkbtnBackToProducts_Click" SkinID="button">&nbsp;Back&nbsp;to&nbsp;Products&nbsp;</asp:LinkButton>
                        &nbsp;
                        <asp:LinkButton ID="lnkbtnProceedToCheckout" runat="server"
                        OnClick="lnkbtnProceedToCheckout_Click" SkinID="button">&nbsp;Proceed&nbsp;to&nbsp;Checkout&nbsp;</asp:LinkButton>
                        </asp:TableCell>
                </asp:TableRow>
            </asp:Table>

           </asp:Panel>
             <asp:Panel id="pnlEmptyBasket" runat="server" visible="False">
           </asp:Panel>
           
            <asp:Panel id="pnlAddress" runat="server">
              <br />
              <table width="90%">
                  <tr>
                      <td style="width: 20%">
                          <strong>
                          Post Code</strong></td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbPostCode" runat="server" Width="400px"></asp:TextBox>
                          <asp:LinkButton ID="lnkbtnFindAddress" runat="server" OnClick="lnkbtnFindAddress_Click">Find&nbsp;Address</asp:LinkButton>&nbsp;
                          <asp:Label ID="lblLookupError" runat="server" Visible="False" ForeColor="Red"></asp:Label></td>
                      <td align="left" style="width: 41%">
                          </td>
                  </tr>
                  <tr id="trPostcodeLookupResults" runat="server" visible="false">
                      <td style="width: 20%">
                          <asp:Label ID="lblSelectADestination" runat="server" Text="Select a destination"></asp:Label></td>
                      <td style="width: 65%">
                          <br />
                          <asp:ListBox ID="lbLookupResults" runat="server" AutoPostBack="True" Width="408px" OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged" Height="250px"></asp:ListBox>
                          <asp:LinkButton ID="lnkbtnAddrLookupCancel" runat="server" OnClick="lnkbtnAddrLookupCancel_Click">Cancel</asp:LinkButton></td>
                      <td align="left" style="width: 41%">
                          </td>
                  </tr>
                  <tr id="trAddr1" runat="server" visible="true">
                      <td style="width: 20%">
                          Attention of:</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbAttnOf" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                          </td>
                  </tr>
                  <tr id="trAddr2" runat="server" visible="false">
                      <td style="width: 20%; height: 26px;">
                          Name <span style="color: red">*</span></td>
                      <td style="width: 65%; height: 26px;">
                          <asp:TextBox ID="tbName" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%; height: 26px" align="left">
                          <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="tbName"
                              ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                  </tr>
                  <tr id="Tr2" runat="server" visible="false">
                      <td style="width: 20%; height: 26px;">
                          Company <span style="color: red">*</span></td>
                      <td style="width: 65%; height: 26px;">
                          <asp:TextBox ID="tbCompany" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 15%; height: 26px" align="left">
                          <asp:RequiredFieldValidator ID="rfvCompany" runat="server" ControlToValidate="tbCompany"
                              ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                  </tr>
                  <tr id="trAddr3" runat="server" visible="true">
                      <td style="width: 20%">
                          Addr 1<span style="color: red">*</span></td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbAddr1" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                          <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="tbAddr1"
                              ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                  </tr>
                  <tr id="trAddr4" runat="server" visible="true">
                      <td style="width: 20%">
                          Addr 2</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbAddr2" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                      </td>
                  </tr>
                  <tr id="trAddr5" runat="server" visible="true">
                      <td style="width: 20%">
                          Addr 3</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbAddr3" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                      </td>
                  </tr>
                  <tr id="trAddr6" runat="server" visible="true">
                      <td style="width: 20%">
                          Town <span style="color: red">*</span></td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbTown" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                          <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="tbTown"
                              ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                  </tr>
                  <tr id="trAddr7" runat="server" visible="true">
                      <td style="width: 20%">
                          County/State</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbCounty" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                      </td>
                  </tr>
                  <tr id="trAddr8" runat="server" visible="true">
                      <td style="width: 20%">
                          Tel.</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbCtcTelNo" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                      </td>
                  </tr>
                  <tr id="Tr1" runat="server" visible="false">
                      <td style="width: 20%">
                          Email</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbCtcEmailAddr" runat="server" Width="400px"></asp:TextBox></td>
                      <td style="width: 15%" align="left">
                          <asp:RequiredFieldValidator ID="rfvCtcEmailAddr" runat="server" ControlToValidate="tbCtcEmailAddr"
                          ErrorMessage="required!" ValidationGroup="Address"></asp:RequiredFieldValidator>
                      </td>
                  </tr>
                  <tr id="trAddr9" runat="server" visible="true">
                      <td style="width: 20%">
                          Country <span style="color: red">*</span></td>
                      <td style="width: 65%">
                          <asp:DropDownList ID="ddlCountry" runat="server" DataSourceID="SqlDataSourceCountries" DataTextField="CountryName" DataValueField="CountryKey">
                          </asp:DropDownList>
                          <asp:SqlDataSource ID="SqlDataSourceCountries" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
                              SelectCommand="spASPNET_Country_GetCountries" SelectCommandType="StoredProcedure"></asp:SqlDataSource>
                          </td>
                      <td style="width: 41%" align="left">
                          <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="ddlCountry"
                              ErrorMessage="please select!" InitialValue="0" ValidationGroup="Address"></asp:RequiredFieldValidator></td>
                  </tr>
                  <tr id="trAddr10" runat="server" visible="true">
                      <td style="width: 20%">
                          Special Instructions<br />(optional)</td>
                      <td style="width: 65%">
                          <asp:TextBox ID="tbSpclInstructions" runat="server" TextMode="MultiLine" Width="400px"></asp:TextBox></td>
                      <td style="width: 41%" align="left">
                      </td>
                  </tr>
              </table>
              <br />
                    <table width="90%">
                        <tr>
                            <td style="width: 100px">
                                <asp:Label ID="lblAddressValidation" runat="server" Text="* Required" ForeColor="Red"></asp:Label>
                            </td>
                            <td align="right">
                        <asp:LinkButton ID="lnkbtnBackToCategoriesFromAddressPanel"
                                        runat="server"
                                        OnClick="lnkbtnProductsByCategory_Click"
                                        SkinID="button">&nbsp;Back&nbsp;to&nbsp;Categories&nbsp;</asp:LinkButton>
                                        &nbsp;
                        <asp:LinkButton ID="lnkbtnBackToProductsFromAddressPanel"
                                        runat="server"
                                        OnClick="lnkbtnBackToProducts_Click"
                                        SkinID="button">&nbsp;Back&nbsp;to&nbsp;Products&nbsp;</asp:LinkButton>
                                        &nbsp;
                       <asp:LinkButton ID="lnkbtnAddressBook"
                                       runat="server"
                                       OnClick="btnAddressBook_Click"
                                       SkinID="button">&nbsp;Address&nbsp;Book&nbsp;</asp:LinkButton>
                                       &nbsp;
                       <asp:LinkButton ID="lnkbtnSubmitOrder"
                                       runat="server"
                                       OnClick="btnSubmitOrder_Click" 
                                       SkinID="button">&nbsp;Submit&nbsp;Order&nbsp;</asp:LinkButton>
                            </td>
                        </tr>
                    </table>
              </asp:Panel>

            <asp:Panel id="pnlAddressList" runat="server" DefaultButton="lnkbtnSearchAddressList">
            <br />
                <asp:SqlDataSource ID="SqlDataSourceAddressList"
                                   runat="server"
                                   ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
                                   ProviderName="System.Data.SqlClient"
                                   SelectCommand="spASPNET_Address_GetGlobalAddresses"
                                   SelectCommandType="StoredProcedure"
                                   >
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
                            Search: <asp:TextBox ID="tbSearchAddressList" runat="server" Text=""></asp:TextBox>&nbsp;<asp:LinkButton ID="lnkbtnSearchAddressList" runat="server" SkinID="button" OnClick="lnkbtnSearchAddressList_Click">Go</asp:LinkButton></td>
                    </tr>
                </table>
                <br />
                <asp:GridView ID="gvAddressList"
                              runat="server"
                              DataSourceID="SqlDataSourceAddressList"
                              AllowPaging="True"
                              AllowSorting="True"
                              PageSize="8"
                              GridLines="None"
                              PagerSettings-Mode="Numeric"
                              AutoGenerateColumns="False"
                              PagerStyle-VerticalAlign="NotSet"
                              PagerStyle-HorizontalAlign="Center" CellPadding="2">
                    <Columns>
                        <asp:TemplateField HeaderText="Name"
                                           SortExpression="Company">
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkbtnSelectedAddress" runat="server" Text='<%# Bind("Company") %>' OnClick="lnkbtnSelectedAddress_Click">LinkButton</asp:LinkButton>
                                <asp:HiddenField ID="hidCountryCode" runat="server" Value='<%# Bind("CountryKey") %>'/>
                                <asp:HiddenField ID="hidDefaultSpecialInstructions" runat="server" Value='<%# Bind("DefaultSpecialInstructions") %>'/>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Addr 1"                        
                                           SortExpression="Addr1">
                            <ItemTemplate>
                                <asp:Label ID="lblAddr1" runat="server" Text='<%# Bind("Addr1") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Addr 2"                        
                                           SortExpression="Addr2">
                            <ItemTemplate>
                                <asp:Label ID="lblAddr2" runat="server" Text='<%# Bind("Addr2") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Addr 3"                        
                                           SortExpression="Addr3">
                            <ItemTemplate>
                                <asp:Label ID="lblAddr3" runat="server" Text='<%# Bind("Addr3") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Town"                        
                                           SortExpression="Town">
                            <ItemTemplate>
                                <asp:Label ID="lblTown" runat="server" Text='<%# Bind("Town") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="County/State"                        
                                           SortExpression="State" Visible="False">
                            <ItemTemplate>
                                <asp:Label ID="lblState" runat="server" Text='<%# Bind("State") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Post Code"                        
                                           SortExpression="PostCode">
                            <ItemTemplate>
                                <asp:Label ID="lblPostCode" runat="server" Text='<%# Bind("PostCode") %>'></asp:Label>
                            </ItemTemplate>
                            <ItemStyle BorderColor="#DDDDDD" BorderStyle="Dotted" BorderWidth="1px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Country"                        
                                           SortExpression="CountryName">
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
                        <asp:LinkButton ID="lnkbtnBackToDeliveryAddress"
                                        runat="server"
                                        OnClick="lnkbtnBackToDeliveryAddress_Click"
                                        SkinID="button">&nbsp;Back&nbsp;to&nbsp;Delivery&nbsp;Address&nbsp;</asp:LinkButton>
                                        &nbsp;
                            </td>
                        </tr>
                    </table>
           </asp:Panel>

            <asp:Panel id="pnlSearch" runat="server" DefaultButton="lnkbtnSearch">
              <table width="70%">
                  <tr>
                      <td >
                          <br />
                          Search for:</td>
                      <td >
                           <br />
                          <asp:TextBox ID="tbSearch" runat="server" MaxLength="30"></asp:TextBox>
                          &nbsp;
                             <asp:LinkButton ID="lnkbtnSearch" runat="server"
                             OnClick="btnSearch_Click"
                             SkinID="button">&nbsp;Go&nbsp;</asp:LinkButton>
                      </td>
                      <td >
                      </td>
                  </tr>
              </table>
              </asp:Panel>

            <asp:Panel id="pnlTrackAndTrace" runat="server" DefaultButton="lnkbtnCheckConsignment">
              <table width="80%">
                  <tr>
                      <td >
                          <br />
                          Consignment No:</td>
                      <td >
                          <br />
                          <asp:TextBox ID="tbConsignmentNo" runat="server" EnableViewState="False" MaxLength="9"></asp:TextBox>
                          &nbsp;
                             <asp:LinkButton ID="lnkbtnCheckConsignment"
                                             runat="server"
                             OnClick="btnCheckConsignment_Click"
                             SkinID="button"
                             >&nbsp;Go&nbsp;</asp:LinkButton>
                      <td >
                      </td>
                  </tr>
                  <tr>
                      <td>
                      </td>
                      <td>
                      <asp:RequiredFieldValidator ID="rfvConsignmentNo" runat="server" ErrorMessage="required!" ControlToValidate="tbConsignmentNo" EnableClientScript="False" ValidationGroup="tracking"></asp:RequiredFieldValidator>
                      <asp:RegularExpressionValidator ID="revConsignmentNo" runat="server" ErrorMessage="must be numeric!" ControlToValidate="tbConsignmentNo" ValidationExpression="^(\d)*$" EnableClientScript="False" ValidationGroup="tracking"></asp:RegularExpressionValidator></td>
                  </tr>
              </table>
              </asp:Panel>
             
            <asp:Panel id="pnlTrackingResult" runat="server">
              <%=GetTracking() %>
              <asp:Label ID="lblTrackingMessage" runat="server"></asp:Label>
              </asp:Panel>

            <asp:Panel id="pnlBookingConfirmation" runat="server">
                <p>
                   Thank you, your order is now being processed.
                </p>
                <p>
                   Please record this consignment number: <span class="CS_RedText"><asp:Label id="lblConsignmentKey" runat="server"></asp:Label></span>.
                </p>
                <p>
                   You can track this consignment using the 'Track and Trace' menu on the left of
                      this page.
                </p>
              </asp:Panel>
            <asp:Panel id="pnlHelp" runat="server">
              </asp:Panel>

  </td>
        </tr>
      </table>
    <asp:Label ID="lblError" runat="server" Text=""></asp:Label>
</asp:Content>
