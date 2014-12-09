<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ import Namespace="Microsoft.Win32" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="FCKeditorV2" Namespace="FredCK.FCKeditorV2" Assembly="FredCK.FCKeditorV2" %>
<script runat="server">

    ' SAW THIS...
    ' Line 9746:                    &nbsp;<telerik:RadNumericTextBox     Cannot create an object of type 'System.Type' from its string representation 'System.Int32' for the 'DataType' property
    
    ' TO DO ON PRODUCT CREDITS
    
    ' add overdraft / enforced to report
    ' report layout improvements
    ' put in commented out script manager
    
    'spASPNET_Product_GetProductFromKey9
    'spASPNET_Product_FullUpdate10
    'spASPNET_Product_AddWithAccessControl9

    ' TO DO
   
    ' ConfigLib.GetConfigItem_EnableRotation - sort this out
    ' ConfigLib.GetConfigItem_CategoryCount - sort this out
   
    ' make SellingPrice an option (shares line with Product Groups, currently set invisible)
   
    ' spASPNET_Customer_UsesCategories_Get
    ' spASPNET_Customer_UsesCategories_Set
    ' spASPNET_LogisticProduct_GetCategoriesCountForCustomer
    ' spASPNET_Product_GetCategories
    ' spASPNET_Product_GetSubCategories
    ' spASPNET_Product_GetSubSubCategories2
    ' spASPNET_Product_GetCustProdsToManageOwned
    ' spASPNET_Product_GetCustProdsToManage3
    ' spASPNET_Product_SetImageAttributes
    ' spASPNET_Product_SetPDFAttribute
    ' spASPNET_Product_GetProductFromKey9
    ' spASPNET_Product_GetAuthorisable2
    ' spASPNET_Product_FullUpdate10
    ' spASPNET_Product_AddWithAccessControl9
    ' spASPNET_Customer_ExplicitProductPermissions_GetFlag
    ' spASPNET_Product_Delete
    ' spASPNET_Product_GetNumbersOwned
    ' spASPNET_Product_GetNumbers
    ' spASPNET_Product_GetUserProfilesFromKey
    ' spASPNET_Product_GetUserProfilesMatchingSearch
    ' spASPNET_Product_SetUserProductProfile
    ' spASPNET_Product_GetPendingAuthorisationRequests
    ' spASPNET_Product_SetAuthorisation
    ' spASPNET_UserProfile_GetAllSuperUsersForCustomer2
    ' spASPNET_Product_SetAuthorisable
    ' spASPNET_Product_RemoveAuthorisable
    ' spASPNET_Product_GetAllPreAuthorisationUsers2
    ' spASPNET_Product_GetPreAuthorisationUsersMatchingSearch2
    ' spASPNET_Product_SetPreAuthorisation
    ' spASPNET_Product_RemoveRotationProduct
    ' spASPNET_Product_GetRotationProducts
    ' spASPNET_Product_SetRotationProduct
    ' spASPNET_Product_CreateProductGroup
    ' spASPNET_Product_RenameProductGroup
    ' spASPNET_Product_UpdatePrimaryProductGroupOwner
    ' spASPNET_Product_UpdateDeputyProductGroupOwner
    ' spASPNET_StockBooking_AuthOrderGetPendingRequests
    ' spASPNET_StockBooking_AuthOrderGetByKey' spASPNET_StockBooking_AuthOrderGetDetails
    ' spASPNET_StockBooking_AuthOrderEmailOrderer
    ' spASPNET_StockBooking_AuthOrderUpdateHoldingQueue
    ' spASPNET_StockBooking_Add3
    ' spASPNET_LogisticMovement_Add
    ' spASPNET_StockBooking_AuthOrderUpdateItemHoldingQueue
    ' spASPNET_LogisticBooking_Complete
    ' spASPNET_Product_GetAuthorised
    ' spASPNET_Product_UpdateAuthoriser
  
    ' spASPNET_Product_GetCategoriesIncludeArchivedProds
    ' spASPNET_Product_GetSubCategoriesIncludeArchivedProds
    ' spASPNET_Product_GetSubSubCategoriesIncludeArchivedProds

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gbDataBound As Boolean = False
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Private gbExplicitProductPermissions As Boolean
    Private bRequiresAuthEnabled As Boolean

    Const NEW_CATEGORY As String = "- new category -"
    Const NEW_SUBCATEGORY As String = "- new subcategory -"
  
    Const DEMO_BAY_KEY As Int32 = 409
    
    Const PER_CUSTOMER_CONFIGURATION_NONE As Integer = 0
    Const PER_CUSTOMER_CONFIGURATION_1_BLACKROCK As Integer = 1
    Const PER_CUSTOMER_CONFIGURATION_5_VSOE As Integer = 5

    Const CUSTOMER_WURS As Integer = 579
    Const CUSTOMER_WURS_TEST_ACCOUNT As Integer = 585

    Const CUSTOMER_WESTERN_UNION As Integer = 651
    Const CUSTOMER_AAT As Integer = 654
    Const CUSTOMER_LOVELLS As Integer = 663
    Const CUSTOMER_QUANTUMLEAP As Int32 = 774
    Const CUSTOMER_BOULEVARD As Int32 = 785
    Const CUSTOMER_JUPITER As Int32 = 784

    Const PER_USERTYPE_OWNER_GROUP_NONE As Integer = 0

    Const CATEGORY_MODE_0_CATEGORIES As Integer = 0
    Const CATEGORY_MODE_1_CATEGORY As Integer = 1
    Const CATEGORY_MODE_2_CATEGORIES As Integer = 2
    Const CATEGORY_MODE_3_CATEGORIES As Integer = 3
  
    Const DISPLAY_MODE_CATEGORY As String = "category"
    Const DISPLAY_MODE_ALL As String = "all"
    Const DISPLAY_MODE_SEARCH As String = "search"

    Const CREDIT_LIMIT_ENFORCE_FALSE As Int32 = 0
    Const CREDIT_LIMIT_ENFORCE_TRUE As Int32 = 1

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call GetSiteFeatures()
            psProdImageFolder = ConfigLib.GetConfigItem_prod_image_folder
            psProdThumbFolder = ConfigLib.GetConfigItem_prod_thumb_folder
            psProdPDFFolder = ConfigLib.GetConfigItem_prod_pdf_folder
  
            psVirtualJPGFolder = ConfigLib.GetConfigItem_Virtual_JPG_URL
            psVirtualThumbFolder = ConfigLib.GetConfigItem_Virtual_Thumb_URL
            psVirtualPDFFolder = ConfigLib.GetConfigItem_Virtual_PDF_URL

            Call PerCustomerConfiguration()

            'If IsRioTinto() Then
            ' Call Set2CategoryLevels()
            'End If

            Call PerUserTypeConfiguration()
            Call GetProductNumbers()
            Call InitProductGroupControls()
            Call ShowMainPanel()

            tbRenamedProductGroup.Attributes.Add("onkeypress", "return clickButton(event,'" + btnRenameThisProductGroup.ClientID + "')")
            tbNewProductGroup.Attributes.Add("onkeypress", "return clickButton(event,'" + btnCreateNewProductGroup.ClientID + "')")
        End If
      
        bRequiresAuthEnabled = pbOrderAuthorisation Or pbProductAuthorisation
        Call SetRequiresAuthAvailability()
      
        SqlDataSourceCategoryList.ConnectionString = ConfigLib.GetConfigItem_ConnectionString

        txtSearchCriteriaAllProducts.Attributes.Add("onkeypress", "return clickButton(event,'" + btn_SearchAllProducts.ClientID + "')")
       
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)
        Call CheckVisibility()
        Response.Buffer = True
        Call SetTitle()
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------- code amended
        If (Not IsPostBack) Then
            btnShowCategories_Click(btnShowCategories, EventArgs.Empty)
            'btn_ShowAllProducts_Click(btn_ShowAllProducts, EventArgs.Empty)

        End If
                     
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------- code amended                    
    End Sub
  
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
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
        Page.Header.Title = sTitle & "Product Manager"
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
            WebMsgBox.Show("GetSiteFeatures: " & ex.Message)
        Finally
            oConn.Close()
        End Try

        Dim dr As DataRow = oDataTable.Rows(0)
        pbOrderAuthorisation = dr("OrderAuthorisation")
        pbProductAuthorisation = dr("ProductAuthorisation")
        pbCalendarManagement = dr("CalendarManagement")
        pbProductOwners = dr("ProductOwners")
        pbProductCredits = dr("ProductCredits")
        pbCustomLetters = dr("CustomLetters")
        pbSellingPrice = dr("SellingPrice")

        pnCategoryMode = dr("CategoryCount")
        If pnCategoryMode = 2 Then
            Call Set2CategoryLevels()
        ElseIf pnCategoryMode = 3 Then
            Call Set3CategoryLevels()
        Else
            WebMsgBox.Show("Error retrieving category levels - please report this problem to your Account Handler")
        End If
    End Sub

    Protected Sub PopulateProductGroupDropdown()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sSQL = "SELECT * FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & " AND (ProductOwner1 = " & Session("UserKey") & " OR ProductOwner2 = " & Session("UserKey") & ") ORDER BY ProductGroupName"
        Else
            sSQL = "SELECT * FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductGroupName"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)

        ddlProductGroup.Items.Clear()
        ddlAssignedProductGroup.Items.Clear()
        ddlProductGroup.Items.Add(New ListItem("- please select -", 0))
        ddlAssignedProductGroup.Items.Add(New ListItem("- please select -", 0))

        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            While oDataReader.Read()
                Dim li As New ListItem
                li.Text = oDataReader("ProductGroupName")
                li.Value = oDataReader("ProductGroupKey")
                ddlProductGroup.Items.Add(li)
                ddlAssignedProductGroup.Items.Add(li)
            End While
        End If
        oConn.Close()
    End Sub
  
    Protected Function GetUsesCategories() As Boolean
        GetUsesCategories = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_UsesCategories_Get", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.Fill(oDataTable)
        If oDataTable.Rows.Count > 0 Then
            If IsDBNull(oDataTable.Rows(0).Item(0)) Then
                GetUsesCategories = False
            Else
                GetUsesCategories = oDataTable.Rows(0).Item(0)
            End If
        End If
    End Function
  
    Protected Sub SetUsesCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Customer_UsesCategories_Set", oConn)
        Dim spParam As SqlParameter
        oCmd.CommandType = CommandType.StoredProcedure

        spParam = New SqlParameter("@CustomerKey", SqlDbType.Int)
        spParam.Value = Session("CustomerKey")
        oCmd.Parameters.Add(spParam)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in SetUsesCategories: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Function CountCategories() As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_LogisticProduct_GetCategoriesCountForCustomer", oConn)

        CountCategories = 0
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.Fill(oDataTable)
        If oDataTable.Rows.Count > 0 Then
            CountCategories = CInt(oDataTable.Rows(0).Item(0))
        End If
    End Function
  
    Protected Sub QueryUsesCategories()
        pbUsesCategories = False
        If GetUsesCategories() = True Then
            pbUsesCategories = True
        Else
            If CountCategories() > 0 Then
                Call SetUsesCategories()
                pbUsesCategories = True
            End If
        End If
    End Sub
  
    Protected Function IsBlackRock() As Boolean
        Dim arrCustomerBlackRock() As Integer = {23}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsBlackRock = IIf(gbSiteTypeDefined, gsSiteType = "blackrock", Array.IndexOf(arrCustomerBlackRock, nCustomerKey) >= 0)
    End Function

    Protected Function IsWURS() As Boolean
        Dim arrCustomerWURS() As Integer = {CUSTOMER_WURS}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWURS = IIf(gbSiteTypeDefined, gsSiteType = "wurs", Array.IndexOf(arrCustomerWURS, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsWesternUnion() As Boolean
        Dim arrCustomerWesternUnion() As Integer = {CUSTOMER_WESTERN_UNION}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWesternUnion = IIf(gbSiteTypeDefined, gsSiteType = "westernunion", Array.IndexOf(arrCustomerWesternUnion, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsVSOE() As Boolean
        'Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        'IsVSOE = IIf(gbSiteTypeDefined, gsSiteType = "vsoe", nCustomerKey = 24)
        Dim arrCustomerVSOE() As Integer = {24, 688}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsVSOE = IIf(gbSiteTypeDefined, gsSiteType = "vsoe", Array.IndexOf(arrCustomerVSOE, nCustomerKey) >= 0)
    End Function
  
    Protected Function IsDAT() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsDAT = IIf(gbSiteTypeDefined, gsSiteType = "dat", nCustomerKey = 546)
    End Function
  
    Protected Function IsHysterOrYale() As Boolean
        Dim arrCustomerHysterYale() As Integer = {77, 680}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHysterOrYale = IIf(gbSiteTypeDefined, gsSiteType = "hysteryale", Array.IndexOf(arrCustomerHysterYale, nCustomerKey) >= 0)
    End Function
  
    Protected Function IsRioTinto() As Boolean
        Dim arrCustomerRioTinto() As Integer = {47, 54, 109}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsRioTinto = IIf(gbSiteTypeDefined, gsSiteType = "riotinto", Array.IndexOf(arrCustomerRioTinto, nCustomerKey) >= 0)
    End Function
  
    Protected Function IsAAT() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsAAT = IIf(gbSiteTypeDefined, gsSiteType = "aat", nCustomerKey = CUSTOMER_AAT)
    End Function
    
    Protected Function IsQuantumLeap() As Boolean
        'Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        'IsQuantumLeap = IIf(gbSiteTypeDefined, gsSiteType = "quantumleap", nCustomerKey = CUSTOMER_QUANTUMLEAP)
        Dim arrCustomer() As Integer = {CUSTOMER_QUANTUMLEAP, CUSTOMER_BOULEVARD}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsQuantumLeap = IIf(gbSiteTypeDefined, gsSiteType = "quantum", Array.IndexOf(arrCustomer, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsJupiter() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsJupiter = IIf(gbSiteTypeDefined, gsSiteType = "jupiter", nCustomerKey = CUSTOMER_JUPITER)
    End Function

    Protected Sub PerUserTypeConfiguration()
        btnProductGroups.Visible = False
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            plOwnerGroup = 0
        Else
            If pbProductOwners Then
                btnProductGroups.Visible = True
            End If
            plOwnerGroup = PER_USERTYPE_OWNER_GROUP_NONE
        End If
    End Sub
  
    Protected Sub PerCustomerConfiguration()
        plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_NONE

        If IsJupiter() Then
            trPrintProperties.Visible = True
            lblLegendMisc2.Visible = False
            txtMisc2.Visible = False
            aHelpMisc2.Visible = False
            aHelpMisc2.InnerText = "NULL"

            ' CN POD 20JUN13
            'ddlPrintType.Items.Clear()
            'ddlPrintType.Items.Add(New ListItem("- please select -", 0))
            'Dim dtPrintType As DataTable = ExecuteQueryToDataTable("SELECT [id], PrintType FROM ClientData_Jupiter_PrintCost ORDER BY [id]")
            'For Each dr As DataRow In dtPrintType.Rows
            '    ddlPrintType.Items.Add(New ListItem(dr("PrintType"), dr("id")))
            'Next
        Else
            trPrintProperties.Visible = False
        End If
        If IsBlackRock() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_1_BLACKROCK
            lblLegendCostCentre.Visible = False
            txtDepartment.Visible = False
            aHelpDepartment.Visible = False
            aHelpDepartment.InnerText = "NULL"

            lblLegendMisc1.Visible = False
            txtMisc1.Visible = False
            aHelpMisc1.Visible = False
            aHelpMisc1.InnerText = "NULL"
           
            lblLegendMisc2.Visible = False
            txtMisc2.Visible = False
            aHelpMisc2.Visible = False
            aHelpMisc2.InnerText = "NULL"
           
            lblLegendSerialNumbers.Visible = False
            chkProspectusNumbers.Visible = False
            aHelpProspectusNumbers.Visible = False
            aHelpProspectusNumbers.InnerText = "NULL"
           
            lblLegendViewOnWebForm.Visible = True
            chkViewOnWebForm.Visible = True
        End If
      
        If IsWURS() Then
            revFEXCOCriticalProduct.Enabled = True
            revFEXCOCriticalProduct.EnableClientScript = True
            lblLegendMisc2.Text = "Critical (Y/N/blank):"
            txtMisc2.MaxLength = 1
        End If
        
        If IsHysterOrYale() Then
            lblLegendUnitValue.Text = "Cost Price (€):"
            lblLegendSellingPrice.Text = "Selling Price (€):"
        End If
        
        If IsDAT() Then
            lblLegendMisc1.Text = "Chrg Vol/NFP @"
            lblLegendMisc2.Text = "Chrg PSOs @"
        End If
       
        If IsQuantumLeap() Then
            lblLegendMisc1.Text = "Supplier:"
            aHelpMisc1.Visible = False
            aHelpMisc1.InnerText = "NULL"
           
            lblLegendMisc2.Text = "Boxed to Ship (Y/N):"
            aHelpMisc2.Visible = False
            aHelpMisc2.InnerText = "NULL"
           
        End If
        
        If pbProductOwners Or pbSellingPrice Then
            trOwner.Visible = True
        End If

        If pbProductOwners Then
            btnProductGroups.Visible = True
            rfdProductGroup.Visible = True
            lblLegendProductGroup.Visible = True
            ddlAssignedProductGroup.Visible = True
            lblAssignedProductOwners.Visible = True
            aHelpAssignedProductGroup.Visible = True
        Else
            rfdProductGroup.Visible = False
            lblLegendProductGroup.Visible = False
            ddlAssignedProductGroup.Visible = False
            lblAssignedProductOwners.Visible = False
            aHelpAssignedProductGroup.Visible = False
            aHelpAssignedProductGroup.InnerText = "NULL"
        End If
      
        If pbProductCredits Then
            btnProductCredits.Visible = True
        Else
            btnProductCredits.Visible = False
        End If
        
        If pbCalendarManagement Then
            lblLegendCalendarManaged.Visible = True
            cbCalendarManaged.Visible = True
            aHelpCalendarManaged.Visible = False
        Else
            lblLegendCalendarManaged.Visible = False
            cbCalendarManaged.Visible = False
            aHelpCalendarManaged.InnerText = "NULL"
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_NONE Then
            Call QueryUsesCategories()
            If pbUsesCategories = False Then
                btnShowCategories.Visible = False
            End If
        End If
       
        lblLegendCustomLetter.Visible = pbCustomLetters
        cbCustomLetter.Visible = pbCustomLetters
        lblLegendProductCredits.Visible = pbProductCredits
        cbProductCredits.Visible = pbProductCredits
        'lnkbtnConfigureProductCredits.Visible = pbProductCredits
        lnkbtnConfigureCustomLetter.Visible = pbCustomLetters
        aHelpCustomLetter.Visible = pbCustomLetters
        If Not pbCustomLetters Then
            aHelpCustomLetter.InnerText = "NULL"
        End If
        
        lblLegendSellingPrice.Visible = pbSellingPrice
        tbSellingPrice.Visible = pbSellingPrice
        aHelpSellingPrice.Visible = pbSellingPrice
        If Not pbSellingPrice Then
            aHelpSellingPrice.InnerText = "NULL"
        End If
    End Sub
  
    Protected Sub Set2CategoryLevels()
        pnCategoryMode = CATEGORY_MODE_2_CATEGORIES
        trSubCategory2.Visible = False
        pnlCategorySelection1.Visible = True
    End Sub

    Protected Sub Set3CategoryLevels()
        pnCategoryMode = CATEGORY_MODE_3_CATEGORIES
        trSubCategory2.Visible = True
        pnlCategorySelection2.Visible = True
    End Sub

    Protected Sub SetRequiresAuthAvailability()
        Dim sbTooltipText As New StringBuilder
        If bRequiresAuthEnabled Then
            If pbIsAddingNew Then
                btnPendingAuthorisations.Visible = False
                cbRequiresAuth.Enabled = False
                sbTooltipText.Append("Product Authorisation is disabled until this product has been created. Click the Save button, then Edit the product to set authorisation.")
            Else
                btnPendingAuthorisations.Visible = True
                cbRequiresAuth.Enabled = True
                lnkbtnPreAuthorise.Enabled = True
                sbTooltipText.Append("Product Authorisation gives you fine control over product ordering by users.")
            End If
        Else
            sbTooltipText.Append("Product Authorisation gives you fine control over product ordering by users.  This feature is currently disabled. To enable Product Authorisation contact your Account Handler.")
            btnPendingAuthorisations.Visible = False
            cbRequiresAuth.Enabled = False
            lnkbtnPreAuthorise.Enabled = False
        End If
        tooltipRequiresAuth.Attributes.Add("onmouseover", "return escape('" & sbTooltipText.ToString & "')")
    End Sub
  
    Protected Sub HideAllPanels()
        pnlMainButtonRow.Visible = False
        pnlProductList.Visible = False
        pnlCategorySelection1.Visible = False
        pnlCategorySelection2.Visible = False
        pnlEditProduct.Visible = False
        pnlProductUserProfile.Visible = False
        pnlAuthoriseProduct.Visible = False
        pnlAuthoriseOrder.Visible = False
        pnlMakeAuthorisable.Visible = False
        pnlRemoveAuthorisable.Visible = False
        pnlProductPreAuthorise.Visible = False
        pnlProductGroups.Visible = False
        pnlNewProductGroup.Visible = False
        pnlRenameProductGroup.Visible = False
        pnlShowAuthOrder.Visible = False
        pnlAssociatedProducts.Visible = False
        pnlProductInactivityAlertStatus.Visible = False
        pnlConfigureProductInactivityAlert.Visible = False
        pnlConfigureCustomLetter.Visible = False
        'pnlNewKODDFISProduct.Visible = False
        pnlConfigureProductCredits.Visible = False
        pnlProductCreditsControl.Visible = False
        lblError.Text = ""
    End Sub

    Protected Sub ShowMainPanel()
        Call HideAllPanels()
        pnlMainButtonRow.Visible = True
        ' pnlProductList.Visible = True
    End Sub
  
    Protected Sub ShowProductList()
        ' ----------------------------------------------------------------------------------------------------------------------------------------------------------        amended codes
        'Call HideAllPanels()
        ' ----------------------------------------------------------------------------------------------------------------------------------------------------------        amended codes        
        pnlMainButtonRow.Visible = True
        pnlProductList.Visible = True
    End Sub
  
    Protected Sub btn_ShowAllProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAllProducts()
    End Sub
  
    Protected Sub ShowAllProducts()
        psDisplayMode = DISPLAY_MODE_ALL
        ' ----------------------------------------------------------------------------------------------------------------------------------------------------------        amended codes
        'Call HideAllPanels()
        ' ----------------------------------------------------------------------------------------------------------------------------------------------------------        amended codes        
        pnlMainButtonRow.Visible = True
        pnlProductList.Visible = True
        dg_ProductList.CurrentPageIndex = 0
        txtSearchCriteriaAllProducts.Text = ""
        Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
    End Sub
  
    Protected Sub btn_SearchAllProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SearchAllProducts()
    End Sub
  
    Protected Sub SearchAllProducts()
        psDisplayMode = DISPLAY_MODE_SEARCH
        Call HideAllPanels()
        pnlMainButtonRow.Visible = True
        pnlProductList.Visible = True
        dg_ProductList.CurrentPageIndex = 0
        Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
    End Sub
  
    Protected Sub lnkbtnShowProductsByCategory_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
            psSubCategory = CStr(e.CommandArgument)
            'lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " "
        Else
            psSubSubCategory = CStr(e.CommandArgument)
            'lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " \ " & psSubSubCategory & " "
        End If
        BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
        Call ShowProductList()
    End Sub
  
    Protected Sub ShowCategories()
        lblSubCategoryHeadingA.Text = "Sub Category"
        lblSubCategoryHeadingB.Visible = False

        psDisplayMode = DISPLAY_MODE_CATEGORY
        txtSearchCriteriaAllProducts.Text = ""
        Call HideAllPanels()
        pnlMainButtonRow.Visible = True
        If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
            pnlCategorySelection1.Visible = True
        Else
            pnlCategorySelection2.Visible = True
        End If
        Repeater2.Visible = False
        Repeater2a.Visible = False
        Repeater3a.Visible = False
        Call GetCategories()
    End Sub
  
    Protected Sub ShowProductDetail()
        Call HideAllPanels()
        pnlEditProduct.Visible = True
        Call CheckVisibility()
    End Sub
  
    Protected Sub ShowAuthoriseProductPanel()
        Call HideAllPanels()
        pnlAuthoriseProduct.Visible = True
    End Sub

    Protected Sub ShowAuthoriseOrderPanel()
        Call HideAllPanels()
        tbAuthOrderMessage.Text = String.Empty
        pnlAuthoriseOrder.Visible = True
    End Sub

    Protected Sub ShowProductUserProfile()
        Call HideAllPanels()
        grid_ProductUsers.CurrentPageIndex = 0
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
        If txtProductDate.Text.Trim.Length > 0 Then
            lblUserPermissionsProductCode.Text = txtProductCode.Text & "-" & txtProductDate.Text
        Else
            lblUserPermissionsProductCode.Text = txtProductCode.Text
        End If
        pnlProductUserProfile.Visible = True
    End Sub
  
    Protected Sub ShowNewProduct()
        Call HideAllPanels()
        txtProductCode.Focus()
        If IsVSOE() Then
            chkArchivedFlag.Checked = True
        End If
        pnlEditProduct.Visible = True
    End Sub
  
    Protected Sub ShowMakeAuthorisable()
        Call HideAllPanels()
        pnlMakeAuthorisable.Visible = True
    End Sub
  
    Protected Sub ShowRemoveAuthorisable()
        Call HideAllPanels()
        pnlRemoveAuthorisable.Visible = True
    End Sub

    Protected Sub ShowPreAuthorise()
        Call HideAllPanels()
        pnlProductPreAuthorise.Visible = True
    End Sub

    Protected Sub ShowProductGroupsPanel()
        Call HideAllPanels()
        pnlProductGroups.Visible = True
    End Sub
  
    Protected Sub ShowNewProductGroupPanel()
        Call HideAllPanels()
        pnlNewProductGroup.Visible = True
    End Sub
  
    Protected Sub ShowRenameProductGroupPanel()
        Call HideAllPanels()
        pnlRenameProductGroup.Visible = True
    End Sub

    Protected Sub ShowAuthOrderDetailsPanel()
        Call HideAllPanels()
        pnlShowAuthOrder.Visible = True
    End Sub
  
    Protected Sub ShowAssociatedProductsPanel()
        Call HideAllPanels()
        Call InitAssociatedProductsPanel()
        pnlAssociatedProducts.Visible = True
    End Sub

    Protected Sub ShowConfigureCustomLetterPanel()
        Call HideAllPanels()
        Call GetCustomLetterConfiguration()
        fckedCustomLetterTemplate.Focus()
        pnlConfigureCustomLetter.Visible = True
    End Sub
    
    Protected Sub InitProductGroupControls()
        If ddlProductGroup.Items.Count = 0 Then
            Call PopulateProductGroupDropdown()
        End If
        If ddlPrimaryProductGroupOwner.Items.Count = 0 Then
            ddlPrimaryProductGroupOwner.Items.Clear()
            ddlDeputyProductGroupOwner.Items.Clear()
            ddlPrimaryProductGroupOwner.Items.Add(New ListItem("- please select -", 0))
            ddlDeputyProductGroupOwner.Items.Add(New ListItem("- please select -", 0))
            For Each kvp As KeyValuePair(Of String, Integer) In GetProductOwners()
                Dim li As New ListItem(kvp.Key, kvp.Value)
                ddlPrimaryProductGroupOwner.Items.Add(li)
                ddlDeputyProductGroupOwner.Items.Add(li)
            Next
        End If
    End Sub
  
    Protected Sub btn_ShowAllUsers_click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProductUserSearch.Text = ""
        lblLegendNoMatchingRecords.Visible = False
        grid_ProductUsers.CurrentPageIndex = 0
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
    End Sub
  
    Protected Sub btn_SearchUsers_click(ByVal sender As Object, ByVal e As System.EventArgs)
        grid_ProductUsers.CurrentPageIndex = 0
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
    End Sub
  
    Protected Sub btnUploadImage_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If fuBrowseImageFile.PostedFile.FileName.Trim <> "" Then
            Call SaveImage()
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub btnUploadPDF_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If fuBrowsePDFFile.PostedFile.FileName.Trim <> "" Then
            Call SavePDF()
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub btnShowCategories_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowCategories()
    End Sub

    Protected Sub repeater1_Item_click(ByVal s As Object, ByVal e As RepeaterCommandEventArgs)
        Dim item As RepeaterItem
        For Each item In s.Items
            Dim x As LinkButton = CType(item.Controls(1), LinkButton)
            x.ForeColor = Navy
        Next
        Dim Link As LinkButton = CType(e.CommandSource, LinkButton)
        Link.ForeColor = Blue
    End Sub
  
    Protected Sub lnkbtnShowSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        'psCategory = CStr(e.CommandArgument)
        'Repeater2.Visible = True
        'Repeater2a.Visible = True
        'Repeater3a.Visible = False

        'lblSubCategoryHeadingA.Text = "Sub Category"
        'lblSubCategoryHeadingB.Visible = False

        'Call GetSubCategories()
        psCategory = CStr(e.CommandArgument)
        If HasSubCategories() Then
            
            Repeater2.Visible = True
            Repeater2a.Visible = True
            Repeater3a.Visible = False

            lblSubCategoryHeadingA.Text = "Sub Category"
            lblSubCategoryHeadingB.Visible = False

            Call GetSubCategories()
        Else
            Call BindProductGridDispatcher(CATEGORY_MODE_1_CATEGORY)
            Call ShowProductList()
        End If
    End Sub
  
    Protected Sub lnkbtnShowSubSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psSubCategory = CStr(e.CommandArgument)
        If HasSubSubCategories() Then
            Repeater3a.Visible = True
            lblSubCategoryHeadingA.Text = "Sub Category 1"
            lblSubCategoryHeadingB.Visible = True
            Call GetSubSubCategories()
        Else
            Call BindProductGridDispatcher(CATEGORY_MODE_2_CATEGORIES)
            Call ShowProductList()
        End If
    End Sub
  
    Protected Sub GetCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "Categories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
                    Repeater1.Visible = True
                    Repeater1.DataSource = oDataSet
                    Repeater1.DataBind()
                Else
                    Repeater1a.Visible = True
                    Repeater1a.DataSource = oDataSet
                    Repeater1a.DataBind()
                End If
            Else
                Repeater1.Visible = False
                Repeater1a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Function HasSubCategories() As Boolean
        HasSubCategories = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oDT As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubCategoriesForUser", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDT)
            If oDT.Rows.Count > 0 Then
                HasSubCategories = True
            End If
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub GetSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubCategories", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "SubCategories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
                    Repeater2.Visible = True
                    Repeater2.DataSource = oDataSet
                    Repeater2.DataBind()
                Else
                    Repeater2a.Visible = True
                    Repeater2a.DataSource = oDataSet
                    Repeater2a.DataBind()
                End If
            Else
                Repeater2.Visible = False
                Repeater2a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Function HasSubSubCategories() As Boolean
        HasSubSubCategories = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oDT As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubSubCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ProductCategory").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDT)
            If oDT.Rows.Count > 0 Then
                HasSubSubCategories = True
            End If
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub GetSubSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetSubSubCategories2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@ProductCategory").Value = psCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "SubSubCategories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                Repeater3a.Visible = True
                Repeater3a.DataSource = oDataSet
                Repeater3a.DataBind()
            Else
                Repeater3a.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub DisplayCategories()
        Dim item As RepeaterItem
        For Each item In Repeater1.Items
            Dim x As LinkButton = CType(item.Controls(1), LinkButton)
            x.ForeColor = Navy
        Next
        Repeater2.Visible = False
        Repeater2a.Visible = False
        Repeater3a.Visible = False
        Call ShowCategories()
    End Sub
  
    Protected Function bDatesValid() As Boolean
        bDatesValid = True
        Dim sExpiryDate As String = tbExpiryDate.Text.Trim
        If sExpiryDate <> String.Empty Then
            Try
                Dim dtExpiryDate As DateTime = DateTime.Parse(sExpiryDate)
                If dtExpiryDate.Year < 2000 Or dtExpiryDate.Year > 2019 Then
                    bDatesValid = False
                End If
            Catch ex As Exception
                bDatesValid = False
            End Try
          
        End If
        Dim sReplenishmentDate As String = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> String.Empty Then
            Try
                Dim dtReplenishmentDate As DateTime = DateTime.Parse(sReplenishmentDate)
                If dtReplenishmentDate.Year < 2000 Or dtReplenishmentDate.Year > 2019 Then
                    bDatesValid = False
                End If
            Catch ex As Exception
                bDatesValid = False
            End Try
          
        End If
    End Function
  
    Protected Sub btn_SaveProductChanges_click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblDateError.Visible = False
        If Page.IsValid Then
            If IsDAT() Then
                txtMisc1.Text = txtMisc1.Text.Trim
                If txtMisc1.Text <> String.Empty Then
                    If Not (IsNumeric(txtMisc1.Text) AndAlso txtMisc1.Text >= 0) Then
                        WebMsgBox.Show("'Charge Vol/NotForProfit @' level must be blank or a positive number, eg 200.\n\nThis is the order level at which charging for this item will begin for Voluntary Organisations, CICs and Not-For-Profit Groups.")
                        Exit Sub
                    End If
                End If
                txtMisc2.Text = txtMisc2.Text.Trim
                If txtMisc2.Text <> String.Empty Then
                    If Not (IsNumeric(txtMisc2.Text) AndAlso txtMisc2.Text >= 0) Then
                        WebMsgBox.Show("'Charge PSO @' level must be blank or a positive number, eg 200.\n\nThis is the order level at which charging for this item will begin for Public Sector Organisations.")
                        Exit Sub
                    End If
                End If
            End If
            If IsJupiter() Then
                Dim nOldProduct As Int32 = 0
                If Not pbIsAddingNew Then
                    nOldProduct = ExecuteQueryToDataTable("SELECT CASE WHEN CreatedOn < '5-jun-2013' THEN 1 ELSE 0 END AS 'old' FROM LogisticProduct WHERE LogisticProductKey = " & plProductKey).Rows(0).Item(0)
                End If
                If cbOnDemand.Checked And nOldProduct = 0 Then
                    Dim bProductDateValid As Boolean = True
                    txtProductDate.Text = txtProductDate.Text.Trim
                    If txtProductDate.Text.Length <> 9 Then
                        bProductDateValid = False
                        
                    ElseIf Not IsNumeric(txtProductDate.Text.Substring(0, 2)) Then
                        bProductDateValid = False
                    ElseIf Not IsNumeric(txtProductDate.Text.Substring(5, 4)) Then
                        bProductDateValid = False
                    ElseIf Not "jan.feb.mar.apr.may.jun.jul.aug.sep.oct.nov.dec".Contains(txtProductDate.Text.Substring(2, 3).ToLower) Then
                        bProductDateValid = False
                    End If
                    If Not bProductDateValid Then
                        WebMsgBox.Show("Product Date must be supplied, and in the format ddMMMyyyy (eg 02May2013).\n\nYou entered '" & txtProductDate.Text & "'.")
                        Exit Sub
                    End If
                End If
            End If
            If bDatesValid() Then
                Call CheckVisibility()          ' do this in case add or update errors out before completion
                If Not gbDataBound Then         ' this is a kludge because the DataBinder is not always called
                    Call AdjustddlCategory()
                    Call AdjustddlSubCategory()
                    If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
                        Call AdjustddlSubSubCategory()
                    End If
                End If
      
                If Not pbIsAddingCategory Then
                    txtCategory.Text = ddlCategory.SelectedItem.ToString
                End If
                If Not pbIsAddingSubCategory Then
                    txtSubCategory.Text = ddlSubCategory.SelectedItem.ToString
                End If
                If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
                    If Not pbIsAddingSubSubCategory Then
                        tbSubSubCategory.Text = ddlSubSubCategory.SelectedItem.ToString
                    End If
                End If
          
                If pbIsAddingNew Then
                    txtProductCode.Text = txtProductCode.Text.Trim
                    txtProductDate.Text = txtProductDate.Text.Trim
                    If txtProductCode.Text.Length > 0 Then
                        gbExplicitProductPermissions = bSetExplicitProductPermissionsFlag()
                        If Session("UserType").ToString.ToLower.Contains("owner") AndAlso ddlAssignedProductGroup.SelectedIndex <= 0 Then
                            WebMsgBox.Show("Please select a Product Group for this product")
                        Else
                            Call AddNewProduct()
                            'If IsJupiter() AndAlso ddlPrintType.SelectedIndex > 0 And lblError.Text = String.Empty Then
                            If IsJupiter() AndAlso ddlPODPageCount.SelectedIndex > 0 And lblError.Text = String.Empty Then   ' CN POD 20JUN13
                                Call AddQuantityToJupiterPODProduct(plProductKey)
                                Call PermissionJupiterPODProduct(plProductKey)
                                Call SetJupiterPODProductArchiveFlag(plProductKey, "Y")
                            End If
                            If cbCustomLetter.Checked Then
                                Call ShowConfigureCustomLetterPanel()
                            End If
                            If gbExplicitProductPermissions Then
                                Call ShowProductUserProfile()
                            End If
                            Call GetProductNumbers()
                        End If
                    Else
                        WebMsgBox.Show("Blank product code not allowed.")
                    End If
                Else
                    If Session("UserType").ToString.ToLower.Contains("owner") AndAlso ddlAssignedProductGroup.SelectedIndex <= 0 Then
                        WebMsgBox.Show("Please select a Product Group for this product")
                    Else
                        Call UpdateProduct()
                        ' If IsJupiter() AndAlso ddlPrintType.SelectedIndex > 0 Then
                        If IsJupiter() AndAlso ddlPODPageCount.SelectedIndex > 0 Then       ' CN POD 20JUN13
                            Call AddQuantityToJupiterPODProduct(plProductKey)
                            Call PermissionJupiterPODProduct(plProductKey)
                        End If
                    End If
                End If
            Else
                lblDateError.Visible = True
            End If
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------- code amended
        'btn_ShowAllProducts_Click(btn_ShowAllProducts, EventArgs.Empty)
        btnShowCategories_Click(btnShowCategories, EventArgs.Empty)
                     
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------- code amended                    
        
    End Sub

    Protected Sub btnSetUserProfiles_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductUserProfile()
    End Sub
  
    Protected Sub btn_DeleteProduct_click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CLng(lblProductQuantity.Text) > 0 Then
            WebMsgBox.Show("You cannot delete a product with a positive stock balance. Pick all remaining stock then delete the product.")
        Else
            If System.IO.File.Exists(psProdImageFolder & plProductKey.ToString & ".jpg") Then
                System.IO.File.Delete(psProdImageFolder & plProductKey.ToString & ".jpg")
            End If
            If System.IO.File.Exists(psProdThumbFolder & plProductKey.ToString & ".jpg") Then
                System.IO.File.Delete(psProdThumbFolder & plProductKey.ToString & ".jpg")
            End If
            If System.IO.File.Exists(psProdPDFFolder & plProductKey.ToString & ".pdf") Then
                System.IO.File.Delete(psProdPDFFolder & plProductKey.ToString & ".pdf")
            End If
            Call DeleteProduct()
        End If
    End Sub
  
    Protected Sub btn_GoToProductListPanel_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BackToProductListPanel()
    End Sub
  
    Protected Sub BackToProductListPanel()
        Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
        Call ShowMainPanel()
    End Sub
  
    Protected Sub btn_GoBackToProductDetail_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReturnToProductDetail()
    End Sub

    Protected Sub ReturnToProductDetail()
        Call ShowProductDetail()
    End Sub
  
    Protected Sub dg_ProductList_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "Edit" Then
            pbIsAddingNew = False
            pbIsAddingCategory = False
            pbIsAddingSubCategory = False
            pbIsAddingSubSubCategory = False
            lblImageUploadUnavailable.Visible = False
            lblPDFUploadUnavailable.Visible = False
            fuBrowseImageFile.Visible = True
            fuBrowsePDFFile.Visible = True
            btnUploadImage.Visible = True
            btnUploadPDF.Visible = True
            btnAssociatedProducts.Visible = True
            Call SetRequiresAuthAvailability()
            Dim cell_Product As TableCell = e.Item.Cells(0)
            If IsNumeric(cell_Product.Text) Then
                plProductKey = CLng(cell_Product.Text)
            End If
            Call GetProductFromKey()
            lnkbtnConfigureCustomLetter.Visible = cbCustomLetter.Checked
            btnSetUserProfiles.Visible = True
            btn_DeleteProduct.Visible = True
            lnkbtnNewProductCode.Visible = False
            Call SetHelpStatus()
            Call ShowProductDetail()
        End If
    End Sub
  
    Protected Sub dg_ProductList_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        dg_ProductList.CurrentPageIndex = e.NewPageIndex
        Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
    End Sub
  
    Protected Sub BindProductGridDispatcher(nCategoryMode As Int32)
        If psDisplayMode = DISPLAY_MODE_CATEGORY Then
            Call BindProductGrid(bUseCategories:=True, nCategoryMode:=nCategoryMode)
        Else
            Call BindProductGrid(bUseCategories:=False, nCategoryMode:=nCategoryMode)
        End If
    End Sub
      
    Protected Sub BindProductGrid(ByVal bUseCategories As Boolean, nCategoryMode As Int32)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim sProc As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sProc = "spASPNET_Product_GetCustProdsToManageOwned2"
        Else
            sProc = "spASPNET_Product_GetCustProdsToManage7"
        End If
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        lblProductMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
          
            oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = txtSearchCriteriaAllProducts.Text
          
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GetByCategory", SqlDbType.Bit))
            oAdapter.SelectCommand.Parameters("@GetByCategory").Value = IIf(bUseCategories, 1, 0)

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CategoryMode", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CategoryMode").Value = nCategoryMode
      
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@Category").Value = psCategory

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SubCategory2").Value = psSubSubCategory

            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                dg_ProductList.DataSource = oDataTable                ' when I navigated to last page in product index, then created a new product, error "invalid current page index". Current page index was 1, Page count was 4, start is 0
                dg_ProductList.DataBind()
                dg_ProductList.Visible = True
                'btn_RefreshProductGrid.Visible= True
                If oDataTable.Rows.Count > 8 Then
                    dg_ProductList.PagerStyle.Visible = True
                Else
                    dg_ProductList.PagerStyle.Visible = False
                End If
            Else
                lblProductMessage.Text = "No products found"
                dg_ProductList.Visible = False
                'btn_RefreshProductGrid.Visible= False
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        If txtSearchCriteriaAllProducts.Text <> String.Empty Then
            txtSearchCriteriaAllProducts.Focus()
        End If
    End Sub
  
    Protected Sub SaveImage()
        Dim fi As New System.IO.FileInfo(fuBrowseImageFile.PostedFile.FileName)
        If fi.Extension.ToLower() = ".jpg" Then
            Try
                Dim sTempPath As String = psProdImageFolder & plProductKey.ToString & ".upload.jpg"
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
                fuBrowseImageFile.PostedFile.SaveAs(sTempPath)
                Call SaveResizedImage(fi, 600, 600) 'new
                Call MakeThumbNail(fi, 60, 60)
                Call SetImageAttributes()
                hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & plProductKey.ToString & ".jpg?" & Now.ToString
                hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & plProductKey.ToString & ".jpg"
                imgbtnDeleteImage.Visible = True
                'now delete the original file
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
            Catch ex As Exception
                Response.Write(ex.ToString)
                hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & "blank_thumb.jpg"
                hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & "blank_image.jpg"
            End Try
        Else
            WebMsgBox.Show("Only files with a .JPG extension can be uploaded")
        End If
    End Sub
  
    Protected Sub SavePDF()
        Dim fi As New System.IO.FileInfo(fuBrowsePDFFile.PostedFile.FileName)
        If fi.Extension.ToLower() = ".pdf" Then
            Try
                Dim sTempPath As String = psProdPDFFolder & plProductKey.ToString & ".pdf"
                If System.IO.File.Exists(sTempPath) Then
                    System.IO.File.Delete(sTempPath)
                End If
                fuBrowsePDFFile.PostedFile.SaveAs(psProdPDFFolder & plProductKey.ToString & ".pdf")
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "pdf_logo.gif"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & plProductKey.ToString & ".pdf"
                hlnk_PDF.Target = "_blank"
                imgbtnDeletePDF.Visible = True
                Call SetPDFAttribute()
                If IsJupiter() Then
                    If cbOnDemand.Checked Then
                        btnUploadPDF.Visible = False
                        imgbtnDeletePDF.Visible = False
                        aHelpUploadPDF.Visible = False
                        fuBrowsePDFFile.Visible = False
                    End If
                    ' If ddlPrintType.SelectedIndex > 0 Then
                    If ddlPODPageCount.SelectedIndex > 0 Then      ' CN POD 20JUN13
                        Call RecordJupiterPDFUpload()
                        Call LogJupiterAuditEvent("PDF_UPLOAD", "PRODUCT: " & plProductKey.ToString & ", USER: " & Session("UserKey"))
                        'Call SetJupiterPODProductArchiveFlag(plProductKey, "N")
                    Else
                        WebMsgBox.Show("The Print Type attribute for this product is not currently set./n/nIf you are uploading a PDF for printing, you MUST first set the Print Type, then upload the document again./n/nThis will ensure the printer is notified of your upload.")
                    End If
                End If
            Catch ex As Exception
                Response.Write(ex.ToString)
            End Try
        Else
            WebMsgBox.Show("Only files with a .PDF extension can be uploaded")
        End If
    End Sub
  
    Protected Sub RecordJupiterPDFUpload()
        Dim sSQL As String = "IF EXISTS (SELECT 1 FROM ClientData_Jupiter_PDFUploads WHERE LogisticProductKey = " & plProductKey & ") UPDATE ClientData_Jupiter_PDFUploads SET ReadyToPrint = 0, UploadOn = GETDATE(), UploadBy = " & Session("UserKey") & " WHERE LogisticProductKey = " & plProductKey & " ELSE INSERT INTO ClientData_Jupiter_PDFUploads (LogisticProductKey, ReadyToPrint, UploadOn, UploadBy) VALUES (" & plProductKey.ToString & ", 0, GETDATE(), " & Session("UserKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub

    Protected Sub LogJupiterAuditEvent(ByVal sEventCode As String, ByVal sEventDescription As String, Optional ByVal nProductKey As Int32 = 0, Optional ByVal nConsignmentKey As Int32 = 0)
        Dim sSQL As String = "INSERT INTO ClientData_Jupiter_AuditTrail (EventCode, EventDescription, ProductKey, ConsignmentKey, EventDateTime, EventAuthor) VALUES ('" & sEventCode & "', '" & sEventDescription & "', " & nProductKey & ", " & nConsignmentKey & ", GETDATE(), " & Session("UserKey") & ")"
        Call JupiterNotification(sEventCode, sEventDescription, nConsignmentKey)
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    'Protected Sub JupiterNotification(ByVal sEventCode As String, ByVal sEventDescription As String)
    '    If sEventCode = "EVENT_NOTIFICATION" Then
    '        Exit Sub
    '    End If
    '    Dim sSQL As String = "SELECT EmailAddr FROM ClientData_Jupiter_EventNotification WHERE EventCode = '" & sEventCode & "'"
    '    Dim sbMessage As New StringBuilder
    '    Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
    '    sbMessage.Append("Jupiter Asset Management Event Notification")
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append("Event Code: ")
    '    sbMessage.Append(sEventCode)
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append("Event Description: ")
    '    sbMessage.Append(sEventDescription)
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append("Please do not reply to this email as replies are not monitored.  Thank you.")
    '    sbMessage.Append(Environment.NewLine)
    '    sbMessage.Append("Transworld")
    '    Dim sPlainTextBody As String = sbMessage.ToString
    '    Dim sHTMLBody As String = sbMessage.ToString.Replace(Environment.NewLine, "<br />" & Environment.NewLine)
    '    For Each dr As DataRow In dt.Rows
    '        Call SendMail("JUPITER_EVENT", dr(0), "Jupiter Event Notification - " & sEventCode, sPlainTextBody, sHTMLBody)
    '        Call LogJupiterAuditEvent("EVENT_NOTIFICATION", sEventCode & " to: " & dr(0))
    '    Next
    'End Sub

    Protected Sub JupiterNotification(ByVal sEventCode As String, ByVal sEventDescription As String, ByVal nConsignmentKey As Int32)
        If sEventCode = "EVENT_NOTIFICATION" Then
            Exit Sub
        End If
        Dim sSQL As String = "SELECT EmailAddr FROM ClientData_Jupiter_EventNotification WHERE EventCode = '" & sEventCode & "'"
        Dim sbMessage As New StringBuilder
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        sbMessage.Append("Jupiter Asset Management Event Notification")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Event Code: ")
        sbMessage.Append(sEventCode)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Event Description: ")
        sbMessage.Append(sEventDescription)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Visit http://my.transworld.eu.com/jupiter to view further information.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Please do not reply to this email as replies are not monitored.  Thank you.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Transworld")
        sbMessage.Append(Environment.NewLine)
        Dim sPlainTextBody As String = sbMessage.ToString
        Dim sHTMLBody As String = sbMessage.ToString.Replace(Environment.NewLine, "<br />" & Environment.NewLine)
        For Each dr As DataRow In dt.Rows
            Dim sRecipient As String = dr(0)
            If sRecipient = "#user#" AndAlso nConsignmentKey > 0 Then
                sSQL = "SELECT EmailAddr FROM UserProfile up INNER JOIN Consignment c ON up.[key] = c.UserKey WHERE c.[key] = " & nConsignmentKey
                Dim dtEmailAddr As DataTable = ExecuteQueryToDataTable(sSQL)
                If dtEmailAddr.Rows.Count = 1 Then
                    sRecipient = dtEmailAddr.Rows(0).Item(0)
                End If
            End If
            Call SendMail("JUPITER_EVENT", sRecipient, "Jupiter Event Notification - " & sEventCode, sPlainTextBody, sHTMLBody)
            Call LogJupiterAuditEvent("EVENT_NOTIFICATION", sEventCode & " to: " & dr(0))
        Next
    End Sub

    Protected Sub SendMail(ByVal sType As String, ByVal sRecipient As String, ByVal sSubject As String, ByVal sBodyText As String, ByVal sBodyHTML As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Email_AddToQueue", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oCmd.Parameters.Add(New SqlParameter("@MessageId", SqlDbType.NVarChar, 20))
            oCmd.Parameters("@MessageId").Value = sType
    
            oCmd.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oCmd.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
            oCmd.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int))
            oCmd.Parameters("@StockBookingKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
            oCmd.Parameters("@ConsignmentKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
            oCmd.Parameters("@ProductKey").Value = 0
    
            oCmd.Parameters.Add(New SqlParameter("@To", SqlDbType.NVarChar, 100))
            oCmd.Parameters("@To").Value = sRecipient
    
            oCmd.Parameters.Add(New SqlParameter("@Subject", SqlDbType.NVarChar, 60))
            oCmd.Parameters("@Subject").Value = sSubject
    
            oCmd.Parameters.Add(New SqlParameter("@BodyText", SqlDbType.NText))
            oCmd.Parameters("@BodyText").Value = sBodyText
    
            oCmd.Parameters.Add(New SqlParameter("@BodyHTML", SqlDbType.NText))
            oCmd.Parameters("@BodyHTML").Value = sBodyHTML
    
            oCmd.Parameters.Add(New SqlParameter("@QueuedBy", SqlDbType.Int))
            oCmd.Parameters("@QueuedBy").Value = Session("UserKey")
    
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SendMail: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub MakeThumbNail(ByVal fi As System.IO.FileInfo, ByVal MaxWidth As Double, ByVal MaxHeight As Double)
        Dim OriginalImg As System.Drawing.Image = System.Drawing.Image.FromFile(psProdImageFolder & plProductKey.ToString & ".upload.jpg")
        Dim TheSize As New System.Drawing.Size(OriginalImg.Width, OriginalImg.Height)
  
        Dim sizer As Double = 1
  
        If (MaxWidth > -1 And TheSize.Width > MaxWidth) Or (MaxHeight > -1 And TheSize.Height > MaxHeight) Then
            If MaxWidth > -1 And TheSize.Width > MaxWidth Then
                sizer = MaxWidth / TheSize.Width
                TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
            End If
            If MaxHeight > -1 And TheSize.Height > MaxHeight Then
                sizer = MaxHeight / TheSize.Height
                TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
            End If
        Else
            TheSize.Width = OriginalImg.Width  'Don't try and reduce an image that's already smaller than target size
            TheSize.Height = OriginalImg.Height
        End If
  
        Dim SavePath As String = psProdThumbFolder & plProductKey.ToString & ".jpg" '& F.Name
  
        Dim NewImg As New System.Drawing.Bitmap(OriginalImg, TheSize)
        OriginalImg.Dispose()
  
        If System.IO.File.Exists(SavePath) Then
            System.IO.File.Delete(SavePath)
        End If
        NewImg.Save(SavePath, System.Drawing.Imaging.ImageFormat.Jpeg)
        NewImg.Dispose()
    End Sub
  
    Protected Sub SaveResizedImage(ByVal fi As System.IO.FileInfo, ByVal MaxWidth As Double, ByVal MaxHeight As Double)
        Dim OriginalImg As System.Drawing.Image = System.Drawing.Image.FromFile(psProdImageFolder & plProductKey.ToString & ".upload.jpg")
        Dim TheSize As New System.Drawing.Size(OriginalImg.Width, OriginalImg.Height)
  
        Dim sizer As Double = 1
        Dim sSavePath As String = psProdImageFolder & plProductKey.ToString & ".jpg"
        Try

            If (MaxWidth > -1 And TheSize.Width > MaxWidth) Or (MaxHeight > -1 And TheSize.Height > MaxHeight) Then
                If MaxWidth > -1 And TheSize.Width > MaxWidth Then
                    sizer = MaxWidth / TheSize.Width
                    TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                    TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
                End If
                If MaxHeight > -1 And TheSize.Height > MaxHeight Then
                    sizer = MaxHeight / TheSize.Height
                    TheSize.Width = Convert.ToInt32(TheSize.Width * sizer)
                    TheSize.Height = Convert.ToInt32(TheSize.Height * sizer)
                End If
            Else
                'Don't try and reduce an image that's already smaller than our target size
                TheSize.Width = OriginalImg.Width
                TheSize.Height = OriginalImg.Height
            End If
  
            Dim NewImg As New System.Drawing.Bitmap(OriginalImg, TheSize)
            OriginalImg.Dispose()
  
            If System.IO.File.Exists(sSavePath) Then
                System.IO.File.Delete(sSavePath)
            End If
            NewImg.Save(sSavePath, System.Drawing.Imaging.ImageFormat.Jpeg)
      
            NewImg.Dispose()

        Catch ex As Exception
            WebMsgBox.Show("Unable to resize image. The system may be too busy to allocated the required amount of memory.  Try resizing your image to make it smaller (maximum dimension 600 pixels). ")
        End Try
    End Sub
  
    Protected Sub SetImageAttributes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetImageAttributes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@ThumbNailImage", SqlDbType.NVarChar, 20)
        paramThumbNailImage.Value = plProductKey.ToString & ".jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Dim paramOriginalImage As SqlParameter = New SqlParameter("@OriginalImage", SqlDbType.NVarChar, 20)
        paramOriginalImage.Value = plProductKey.ToString & ".jpg"
        oCmd.Parameters.Add(paramOriginalImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in SetImageAttributes: " & ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub ResetImageAttributes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetImageAttributes", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@ThumbNailImage", SqlDbType.NVarChar, 20)
        paramThumbNailImage.Value = "blank_thumb.jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Dim paramOriginalImage As SqlParameter = New SqlParameter("@OriginalImage", SqlDbType.NVarChar, 20)
        paramOriginalImage.Value = "blank_image.jpg"
        oCmd.Parameters.Add(paramOriginalImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in ResetImageAttributes: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub SetPDFAttribute()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetPDFAttribute", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@PDFFileName", SqlDbType.NVarChar, 60)
        paramThumbNailImage.Value = plProductKey.ToString & ".pdf"
        oCmd.Parameters.Add(paramThumbNailImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in SetPDFAttribute " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub ResetPDFAttribute()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetPDFAttribute", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Dim paramThumbNailImage As SqlParameter = New SqlParameter("@PDFFileName", SqlDbType.NVarChar, 60)
        paramThumbNailImage.Value = "blank_pdf.jpg"
        oCmd.Parameters.Add(paramThumbNailImage)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = "Error in ResetPDFAttribute: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub GetProductFromKey()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetProductFromKey10", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int, 4)
        oParam.Value = plProductKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If IsDBNull(oDataReader("ProductCode")) Then
                txtProductCode.Text = ""
            Else
                txtProductCode.Text = oDataReader("ProductCode")
            End If
            If IsDBNull(oDataReader("ProductDate")) Then
                txtProductDate.Text = ""
            Else
                txtProductDate.Text = oDataReader("ProductDate")
            End If
            If IsDBNull(oDataReader("ProductDescription")) Then
                txtDescription.Text = ""
            Else
                txtDescription.Text = oDataReader("ProductDescription")
            End If
            If IsDBNull(oDataReader("LanguageId")) Then
                txtLanguage.Text = ""
            Else
                txtLanguage.Text = oDataReader("LanguageId")
            End If
            If IsDBNull(oDataReader("ProductDepartmentId")) Then
                txtDepartment.Text = ""
            Else
                txtDepartment.Text = oDataReader("ProductDepartmentId")
            End If
            If IsDBNull(oDataReader("ProductCategory")) Then
                hidCategory.Value = ""
            Else
                hidCategory.Value = oDataReader("ProductCategory")
            End If
            If IsDBNull(oDataReader("SubCategory")) Then
                hidSubCategory.Value = ""
            Else
                hidSubCategory.Value = oDataReader("SubCategory")
            End If
            If IsDBNull(oDataReader("SubCategory2")) Then
                hidSubSubCategory.Value = ""
            Else
                hidSubSubCategory.Value = oDataReader("SubCategory2")
            End If
            If IsDBNull(oDataReader("Notes")) Then
                txtNotes.Text = ""
            Else
                txtNotes.Text = oDataReader("Notes")
            End If
            If IsDBNull(oDataReader("ItemsPerBox")) Then
                txtItemsPerBox.Text = ""
            Else
                txtItemsPerBox.Text = oDataReader("ItemsPerBox")
            End If
            If IsDBNull(oDataReader("MinimumStockLevel")) Then
                txtMinStockLevel.Text = ""
            Else
                txtMinStockLevel.Text = oDataReader("MinimumStockLevel")
            End If
            If IsDBNull(oDataReader("UnitValue")) Then
                txtUnitValue.Text = ""
            Else
                txtUnitValue.Text = Format(oDataReader("UnitValue"), "#,##0.00")
            End If
            If IsDBNull(oDataReader("UnitValue2")) Then
                tbSellingPrice.Text = ""
            Else
                tbSellingPrice.Text = Format(oDataReader("UnitValue2"), "#,##0.00")
            End If
            If IsDBNull(oDataReader("UnitWeightGrams")) Then
                txtUnitWeight.Text = ""
            Else
                txtUnitWeight.Text = Format(oDataReader("UnitWeightGrams"), "#,##0")
            End If
            ' expiry & replenishment date handling note: the extra < 1990 test is to cater for possible 1/1/1900 dates (equivalent to NULL) in the database
            If IsDBNull(oDataReader("ExpiryDate")) Then
                tbExpiryDate.Text = ""
            Else
                If Year(oDataReader("ExpiryDate")) < 1990 Then
                    tbExpiryDate.Text = ""
                Else
                    tbExpiryDate.Text = Format(oDataReader("ExpiryDate"), "dd-MMM-yyyy")
                End If
            End If
            If IsDBNull(oDataReader("ReplenishmentDate")) Then
                tbReplenishmentDate.Text = ""
            Else
                If Year(oDataReader("ReplenishmentDate")) < 1990 Then
                    tbReplenishmentDate.Text = ""
                Else
                    tbReplenishmentDate.Text = Format(oDataReader("ReplenishmentDate"), "dd-MMM-yyyy")
                End If
            End If
            If IsDBNull(oDataReader("Misc1")) Then
                txtMisc1.Text = ""
            Else
                txtMisc1.Text = oDataReader("Misc1")
            End If
            If IsJupiter() Then
                ' ddlPrintType.SelectedIndex = 0
                ''ddlPODPageCount.SelectedIndex = 0      ' CN POD 20JUN13
                ''rfdPrintType.Enabled = False ' ???????????????? CHANGE TO validate page count???????
                ''lblLegendPrintType.Visible = False
                ' ddlPrintType.Visible = False     
                ''ddlPODPageCount.Visible = False         ' CN POD 20JUN13
                ''cbOnDemand.Checked = False
                If Not IsDBNull(oDataReader("Misc2")) Then
                    Call SetPODAttributes(sPODPrintCostIndex:=oDataReader("Misc2"))
                    ''If IsNumeric(oDataReader("Misc2")) Then
                    ''    If oDataReader("Misc2") > 0 Then
                    ''        rfdPrintType.Enabled = True
                    ''        lblLegendPrintType.Visible = True
                    ''        ' ddlPrintType.Visible = True    
                    ''        ' ddlPrintType.SelectedIndex = 0
                    ''        ddlPODPageCount.Visible = True     ' CN POD 20JUN13
                    ''        ddlPODPageCount.SelectedIndex = 0  ' CN POD 20JUN13
                    ''        cbOnDemand.Checked = True
                    ''        For i As Int32 = 0 To ddlPrintType.Items.Count - 1
                    ''            If oDataReader("Misc2") = ddlPrintType.Items(i).Value Then
                    ''                ddlPrintType.SelectedIndex = i
                    ''                Exit For
                    ''            End If
                    ''        Next
                    ''    End If
                    ''End If
                Else
                    Call SetPODAttributes(sPODPrintCostIndex:=String.Empty)
                End If
            Else
                If IsDBNull(oDataReader("Misc2")) Then
                    txtMisc2.Text = ""
                Else
                    txtMisc2.Text = oDataReader("Misc2")
                End If
            End If
          
            ddlAssignedProductGroup.SelectedIndex = 0

            Dim nStockOwnedByKey As Integer = 0
            If Not IsDBNull(oDataReader("StockOwnedByKey")) Then
                nStockOwnedByKey = CInt(oDataReader("StockOwnedByKey"))
            End If
            For i As Integer = 1 To ddlAssignedProductGroup.Items.Count - 1
                If ddlAssignedProductGroup.Items(i).Value = nStockOwnedByKey Then
                    ddlAssignedProductGroup.SelectedIndex = i
                    Exit For
                End If
            Next
            Call SetProductOwners()

            If IsDBNull(oDataReader("SerialNumbersFlag")) Then
                chkProspectusNumbers.Checked = False
            ElseIf oDataReader("SerialNumbersFlag") = "Y" Then
                chkProspectusNumbers.Checked = True
            ElseIf oDataReader("SerialNumbersFlag") = "N" Then
                chkProspectusNumbers.Checked = False
            End If
            If IsDBNull(oDataReader("ArchiveFlag")) Then
                chkArchivedFlag.Checked = False
            ElseIf oDataReader("ArchiveFlag") = "Y" Then
                chkArchivedFlag.Checked = True
            ElseIf oDataReader("ArchiveFlag") = "N" Then
                chkArchivedFlag.Checked = False
            End If
            hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & oDataReader("ThumbNailImage")
            hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & oDataReader("OriginalImage")
            If oDataReader("ThumbNailImage") = "blank_thumb.jpg" Then
                imgbtnDeleteImage.Visible = False
            Else
                imgbtnDeleteImage.Visible = True
            End If
          
            If oDataReader("PDFFileName") = "blank_pdf.jpg" Then
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "blank_pdf_thumb.jpg"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & "blank_pdf.jpg"
                imgbtnDeletePDF.Visible = False
            Else
                hlnk_PDF.ImageUrl = psVirtualPDFFolder & "pdf_logo.gif"
                hlnk_PDF.NavigateUrl = psVirtualPDFFolder & oDataReader("PDFFileName")
                imgbtnDeletePDF.Visible = True
                If cbOnDemand.Checked Then
                    btnUploadPDF.Visible = False
                    imgbtnDeletePDF.Visible = False
                    aHelpUploadPDF.Visible = False
                    fuBrowsePDFFile.Visible = False
                End If
            End If
  
            If oDataReader("WebsiteAdRotatorFlag") = True Then
                chkAdRotator.Checked = True
            Else
                chkAdRotator.Checked = False
            End If
            If IsDBNull(oDataReader("AdRotatorText")) Then
                txtAdRotatorText.Text = ""
            Else
                txtAdRotatorText.Text = oDataReader("AdRotatorText")
            End If
            If oDataReader("ViewOnWebForm") = True Then
                chkViewOnWebForm.Checked = True
            Else
                chkViewOnWebForm.Checked = False
            End If
            If IsDBNull(oDataReader("Flag1")) Then
                cbViewOnWebFormDE.Checked = False
            Else
                If oDataReader("Flag1") = True Then
                    cbViewOnWebFormDE.Checked = True
                Else
                    cbViewOnWebFormDE.Checked = False
                End If
            End If

            If IsDBNull(oDataReader("InactivityAlertDays")) Then
                tbInactivityAlertDays.Text = 0
            Else
                tbInactivityAlertDays.Text = oDataReader("InactivityAlertDays")
            End If
            If IsDBNull(oDataReader("CalendarManaged")) Then
                cbCalendarManaged.Checked = False
            Else
                If oDataReader("CalendarManaged") = True Then
                    cbCalendarManaged.Checked = True
                    lblLegendLanguage.Text = "Type:"
                Else
                    cbCalendarManaged.Checked = False
                    lblLegendLanguage.Text = "Language:"
                End If
            End If
            If IsQuantumLeap() Then
                lblLegendLanguage.Text = "Barcode:"
            End If
           
            lblProductQuantity.Text = Format(oDataReader("Quantity"), "#,##0")
            
            If IsDBNull(oDataReader("CustomLetter")) Then
                cbCustomLetter.Checked = False
            Else
                If oDataReader("CustomLetter") = 0 Then
                    cbCustomLetter.Checked = False
                Else
                    cbCustomLetter.Checked = True
                End If
            End If
            lnkbtnConfigureCustomLetter.Visible = cbCustomLetter.Checked

            If oDataReader("ProductCredits") = 0 Then
                cbProductCredits.Checked = False
                lnkbtnConfigureProductCredits.Visible = False
            Else
                cbProductCredits.Checked = True
                lnkbtnConfigureProductCredits.Visible = True
            End If

            oDataReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call GetRequireAuth()
    End Sub
  
    Protected Sub SetPODAttributes(ByVal sPODPrintCostIndex As String)
        cbOnDemand.Checked = False
        lblPODPrintType.Text = "(undefined)"
        Call SetOnDemandControlsVisibility(False)
        If Not IsNumeric(sPODPrintCostIndex) Then
            sPODPrintCostIndex = String.Empty
        End If
        If Not sPODPrintCostIndex = String.Empty Then
            Dim sSQL As String = "SELECT * FROM ClientData_Jupiter_PrintCost WHERE [id] = " & sPODPrintCostIndex
            Dim dtPODPrintCost As DataTable = ExecuteQueryToDataTable(sSQL)
            If dtPODPrintCost.Rows.Count = 1 Then
                cbOnDemand.Checked = True
                Dim drPODPrintCost As DataRow = dtPODPrintCost.Rows(0)
                Dim nPageCount As Int32 = drPODPrintCost("DocPages")
                Dim nStockSize As Int32 = drPODPrintCost("StockSize")
                Dim nStockWeight As Int32 = drPODPrintCost("StockWeight")
                For i As Int32 = 0 To ddlPODPageCount.Items.Count - 1
                    If nPageCount = ddlPODPageCount.Items(i).Value Then
                        ddlPODPageCount.SelectedIndex = i
                        Exit For
                    End If
                Next
                If nStockSize = 200 Then
                    rbPOD200gsm.Checked = True
                ElseIf nStockSize = 150 Then
                    rbPOD150gsm.Checked = True
                Else
                    rbPOD120gsm.Checked = True
                End If
                If nStockWeight = 5 Then
                    rbPODSizeA5.Checked = True
                Else
                    rbPODSizeA4.Checked = True
                End If
                lblPODPrintType.Text = drPODPrintCost("PrintType")
            End If
            Call SetOnDemandControlsVisibility(True)
        End If
    End Sub
    
    Protected Function GetAuthoriser() As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetAuthorisable2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
      
        oConn.Open()
        Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
        If oSqlDataReader.HasRows Then
            oSqlDataReader.Read()
            GetAuthoriser = oSqlDataReader("FirstName") & " " & oSqlDataReader("LastName")
        Else
            GetAuthoriser = String.Empty
        End If
        oConn.Close()
    End Function

    Protected Sub GetRequireAuth()
        If btnPendingAuthorisations.Visible = False Then
            cbRequiresAuth.Enabled = False
            lnkbtnPreAuthorise.Visible = False
            spanProductAuthorisationAd.Visible = True
            Exit Sub
        End If
      
        Dim sAuthoriser As String = GetAuthoriser()
        If Not sAuthoriser = String.Empty Then
            sAuthoriser = "Authoriser: " & sAuthoriser
            cbRequiresAuth.Checked = True
            tooltipRequiresAuth.Attributes.Add("onmouseover", "return escape('" & sAuthoriser & "')")
            tooltipRequiresAuth.Visible = True
        Else
            cbRequiresAuth.Checked = False
            lnkbtnPreAuthorise.Visible = False
            tooltipRequiresAuth.Visible = False
        End If
    End Sub
  
    Protected Function GetJupiterPODIndex() As Int32
        GetJupiterPODIndex = -1
        If cbOnDemand.Checked Then
            If ddlPODPageCount.SelectedIndex > 0 Then
                Dim sSQL As String = "SELECT ISNULL([id], 0) FROM ClientData_Jupiter_PrintCost WHERE DocPages = " & ddlPODPageCount.SelectedValue & " AND StockSize = "
                If rbPODSizeA4.Checked Then
                    sSQL &= "4"
                Else
                    sSQL &= "5"
                End If
                sSQL &= " AND StockWeight = "
                If rbPOD120gsm.Checked Then
                    sSQL &= "120"
                ElseIf rbPOD150gsm.Checked Then
                    sSQL &= "150"
                Else
                    sSQL &= "200"
                End If
                sSQL &= " AND IsDeleted = 'N'"
                GetJupiterPODIndex = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
            Else
                GetJupiterPODIndex = 0
            End If
        End If
    End Function
    
    Protected Sub UpdateProduct()
        lblError.Text = ""
        Dim nJupiterPODIndex As Int32 = GetJupiterPODIndex()
        If nJupiterPODIndex = 0 Then
            WebMsgBox.Show("This combination of document attributes is not supported.\n\nPlease adjust the attributes.")
            Exit Sub
        End If

        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_FullUpdate11", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        If IsNumeric(txtMinStockLevel.Text) Then
            paramMinimumStockLevel.Value = CLng(txtMinStockLevel.Text)
        Else
            paramMinimumStockLevel.Value = 0
        End If
        oCmd.Parameters.Add(paramMinimumStockLevel)
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = txtDescription.Text
        oCmd.Parameters.Add(paramDescription)
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        If IsNumeric(txtItemsPerBox.Text) Then
            paramItemsPerBox.Value = CLng(txtItemsPerBox.Text)
        Else
            paramItemsPerBox.Value = 0
        End If
        oCmd.Parameters.Add(paramItemsPerBox)
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = txtCategory.Text
        oCmd.Parameters.Add(paramCategory)
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = txtSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory)
        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = tbSubSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory2)
        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        If IsNumeric(txtUnitValue.Text) Then
            If CDec(txtUnitValue.Text) > 0 Then
                paramUnitValue.Value = CDec(txtUnitValue.Text)
            Else
                paramUnitValue.Value = 0
            End If
        Else
            paramUnitValue.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue)
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        If IsNumeric(tbSellingPrice.Text) Then
            If CDec(tbSellingPrice.Text) > 0 Then
                paramUnitValue2.Value = CDec(tbSellingPrice.Text)
            Else
                paramUnitValue2.Value = 0
            End If
        Else
            paramUnitValue2.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue2)
        
        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = txtLanguage.Text
        oCmd.Parameters.Add(paramLanguage)
        
        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = txtDepartment.Text
        oCmd.Parameters.Add(paramDepartment)
        
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        If IsNumeric(txtUnitWeight.Text) Then
            paramWeight.Value = CLng(txtUnitWeight.Text)
        Else
            paramWeight.Value = 0
        End If
        oCmd.Parameters.Add(paramWeight)
      
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        paramStockOwnedByKey.Value = ddlAssignedProductGroup.SelectedValue
        oCmd.Parameters.Add(paramStockOwnedByKey)
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = txtMisc1.Text
        oCmd.Parameters.Add(paramMisc1)
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        If IsJupiter() Then
            If nJupiterPODIndex > 0 Then
                'paramMisc2.Value = ddlPrintType.SelectedValue
                paramMisc2.Value = nJupiterPODIndex.ToString
            Else
                paramMisc2.Value = txtMisc2.Text
            End If
        Else
            paramMisc2.Value = txtMisc2.Text
        End If
        oCmd.Parameters.Add(paramMisc2)
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        If chkArchivedFlag.Checked Then
            paramArchive.Value = "Y"
        Else
            paramArchive.Value = "N"
        End If
        oCmd.Parameters.Add(paramArchive)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        paramStatus.Value = 0
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        Dim sExpiryDate As String
        sExpiryDate = tbExpiryDate.Text.Trim()
        If sExpiryDate <> "" Then
            Try
                sExpiryDate = DateTime.Parse(sExpiryDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Expiry Date"
                Exit Sub
            End Try
        End If
        If sExpiryDate = "" Then
            paramExpiryDate.Value = Nothing
        Else
            paramExpiryDate.Value = sExpiryDate
        End If
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        Dim sReplenishmentDate As String
        sReplenishmentDate = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> "" Then
            Try
                sReplenishmentDate = DateTime.Parse(sReplenishmentDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Renewal / Review Date"
                Exit Sub
            End Try
        End If
        If sReplenishmentDate = "" Then
            paramReplenishmentDate.Value = Nothing
        Else
            paramReplenishmentDate.Value = sReplenishmentDate
        End If
        oCmd.Parameters.Add(paramReplenishmentDate)

        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        If chkProspectusNumbers.Checked Then
            paramSerialNumbers.Value = "Y"
        Else
            paramSerialNumbers.Value = "N"
        End If
        oCmd.Parameters.Add(paramSerialNumbers)
        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = txtAdRotatorText.Text
        oCmd.Parameters.Add(paramAdRotatorText)

        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        If chkAdRotator.Checked Then
            paramWebsiteAdRotatorFlag.Value = 1
        Else
            paramWebsiteAdRotatorFlag.Value = 0
        End If
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)

        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = txtNotes.Text
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWeb As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        If chkViewOnWebForm.Checked Then
            paramViewOnWeb.Value = 1
        Else
            paramViewOnWeb.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWeb)
      
        Dim paramViewOnWebFormDE As SqlParameter = New SqlParameter("@Flag1", SqlDbType.Bit)
        If cbViewOnWebFormDE.Checked Then
            paramViewOnWebFormDE.Value = 1
        Else
            paramViewOnWebFormDE.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWebFormDE)

        Dim paramFlag2 As SqlParameter = New SqlParameter("@Flag2", SqlDbType.Bit)
        paramFlag2.Value = 0
        oCmd.Parameters.Add(paramFlag2)

        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int)
        If IsNumeric(tbInactivityAlertDays.Text) Then
            paramInactivityAlertDays.Value = CLng(tbInactivityAlertDays.Text)
        Else
            paramInactivityAlertDays.Value = 0
        End If
        oCmd.Parameters.Add(paramInactivityAlertDays)
      
        Dim paramCalendarManaged As SqlParameter = New SqlParameter("@CalendarManaged", SqlDbType.Bit)
        If cbCalendarManaged.Checked Then
            paramCalendarManaged.Value = 1
        Else
            paramCalendarManaged.Value = 0
        End If
        oCmd.Parameters.Add(paramCalendarManaged)

        Dim paramOnDemand As SqlParameter = New SqlParameter("@OnDemand", SqlDbType.Int)
        paramOnDemand.Value = 0
        oCmd.Parameters.Add(paramOnDemand)
        
        Dim paramOnDemandPriceList As SqlParameter = New SqlParameter("@OnDemandPriceList", SqlDbType.Int)
        paramOnDemandPriceList.Value = 0
        oCmd.Parameters.Add(paramOnDemandPriceList)
        
        Dim paramCustomLetter As SqlParameter = New SqlParameter("@CustomLetter", SqlDbType.Bit)
        If pbCustomLetters Then
            If cbCustomLetter.Checked Then
                paramCustomLetter.Value = 1
            Else
                paramCustomLetter.Value = 0
            End If
        Else
            paramCustomLetter.Value = 0
        End If
        oCmd.Parameters.Add(paramCustomLetter)
        
        Dim paramProductCredits As SqlParameter = New SqlParameter("@ProductCredits", SqlDbType.Bit)
        If pbProductCredits Then
            If cbProductCredits.Checked Then
                paramProductCredits.Value = 1
            Else
                paramProductCredits.Value = 0
            End If
        Else
            paramProductCredits.Value = 0
        End If
        oCmd.Parameters.Add(paramProductCredits)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
            Call ShowMainPanel()
        Catch ex As SqlException
            lblError.Text = "Error in UpdateProduct: " & ex.Message
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub AddNewProduct()
        lblError.Text = ""
        Dim nJupiterPODIndex As Int32 = GetJupiterPODIndex()
        If nJupiterPODIndex = 0 Then
            WebMsgBox.Show("This combination of document attributes is not supported.\n\nPlease adjust the attributes.")
            Exit Sub
        End If
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddWithAccessControl10", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
  
        Dim paramProductCode As SqlParameter = New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25)
        paramProductCode.Value = txtProductCode.Text
        oCmd.Parameters.Add(paramProductCode)
        Dim paramProductDate As SqlParameter = New SqlParameter("@ProductDate", SqlDbType.NVarChar, 10)
        paramProductDate.Value = txtProductDate.Text
        oCmd.Parameters.Add(paramProductDate)
  
        Dim paramMinimumStockLevel As SqlParameter = New SqlParameter("@MinimumStockLevel", SqlDbType.Int, 4)
        If IsNumeric(txtMinStockLevel.Text) Then
            paramMinimumStockLevel.Value = CLng(txtMinStockLevel.Text)
        Else
            paramMinimumStockLevel.Value = 0
        End If
        oCmd.Parameters.Add(paramMinimumStockLevel)
        Dim paramDescription As SqlParameter = New SqlParameter("@ProductDescription", SqlDbType.NVarChar, 300)
        paramDescription.Value = txtDescription.Text
        oCmd.Parameters.Add(paramDescription)
        Dim paramItemsPerBox As SqlParameter = New SqlParameter("@ItemsPerBox", SqlDbType.Int, 4)
        If IsNumeric(txtItemsPerBox.Text) Then
            paramItemsPerBox.Value = CLng(txtItemsPerBox.Text)
        Else
            paramItemsPerBox.Value = 0
        End If
        oCmd.Parameters.Add(paramItemsPerBox)
        Dim paramCategory As SqlParameter = New SqlParameter("@ProductCategory", SqlDbType.NVarChar, 50)
        paramCategory.Value = txtCategory.Text
        oCmd.Parameters.Add(paramCategory)
        Dim paramSubCategory As SqlParameter = New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50)
        paramSubCategory.Value = txtSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory)
        Dim paramSubCategory2 As SqlParameter = New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50)
        paramSubCategory2.Value = tbSubSubCategory.Text
        oCmd.Parameters.Add(paramSubCategory2)
        Dim paramUnitValue As SqlParameter = New SqlParameter("@UnitValue", SqlDbType.Money, 8)
        If IsNumeric(txtUnitValue.Text) Then
            If CDec(txtUnitValue.Text) > 0 Then
                paramUnitValue.Value = CDec(txtUnitValue.Text)
            Else
                paramUnitValue.Value = 0
            End If
        Else
            paramUnitValue.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue)
        Dim paramUnitValue2 As SqlParameter = New SqlParameter("@UnitValue2", SqlDbType.Money, 8)
        If IsNumeric(tbSellingPrice.Text) Then
            If CDec(tbSellingPrice.Text) > 0 Then
                paramUnitValue2.Value = CDec(tbSellingPrice.Text)
            Else
                paramUnitValue2.Value = 0
            End If
        Else
            paramUnitValue2.Value = 0
        End If
        oCmd.Parameters.Add(paramUnitValue2)

        Dim paramLanguage As SqlParameter = New SqlParameter("@LanguageId", SqlDbType.NVarChar, 20)
        paramLanguage.Value = txtLanguage.Text
        oCmd.Parameters.Add(paramLanguage)

        Dim paramDepartment As SqlParameter = New SqlParameter("@ProductDepartmentId", SqlDbType.NVarChar, 20)
        paramDepartment.Value = txtDepartment.Text
        oCmd.Parameters.Add(paramDepartment)
        Dim paramWeight As SqlParameter = New SqlParameter("@UnitWeightGrams", SqlDbType.Int, 4)
        If IsNumeric(txtUnitWeight.Text) Then
            paramWeight.Value = CLng(txtUnitWeight.Text)
        Else
            paramWeight.Value = 0
        End If
        oCmd.Parameters.Add(paramWeight)
        Dim paramStockOwnedByKey As SqlParameter = New SqlParameter("@StockOwnedByKey", SqlDbType.Int, 4)
        If IsNumeric(ddlAssignedProductGroup.SelectedValue) Then
            paramStockOwnedByKey.Value = ddlAssignedProductGroup.SelectedValue
        Else
            paramStockOwnedByKey.Value = 0
        End If
        oCmd.Parameters.Add(paramStockOwnedByKey)
        Dim paramMisc1 As SqlParameter = New SqlParameter("@Misc1", SqlDbType.NVarChar, 50)
        paramMisc1.Value = txtMisc1.Text
        oCmd.Parameters.Add(paramMisc1)
        Dim paramMisc2 As SqlParameter = New SqlParameter("@Misc2", SqlDbType.NVarChar, 50)
        If nJupiterPODIndex > 0 Then
            paramMisc2.Value = nJupiterPODIndex.ToString
        Else
            paramMisc2.Value = txtMisc2.Text
        End If
        oCmd.Parameters.Add(paramMisc2)
        Dim paramArchive As SqlParameter = New SqlParameter("@ArchiveFlag", SqlDbType.NVarChar, 1)
        If chkArchivedFlag.Checked Then
            paramArchive.Value = "Y"
        Else
            paramArchive.Value = "N"
        End If
        oCmd.Parameters.Add(paramArchive)
      
        Dim paramStatus As SqlParameter = New SqlParameter("@Status", SqlDbType.TinyInt)
        paramStatus.Value = 0
        oCmd.Parameters.Add(paramStatus)

        Dim paramExpiryDate As SqlParameter = New SqlParameter("@ExpiryDate", SqlDbType.SmallDateTime)
        Dim sExpiryDate As String
        sExpiryDate = tbExpiryDate.Text.Trim()
        If sExpiryDate <> "" Then
            Try
                sExpiryDate = DateTime.Parse(sExpiryDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Expiry Date"
                Exit Sub
            End Try
        End If
        If sExpiryDate = "" Then
            paramExpiryDate.Value = Nothing
        Else
            paramExpiryDate.Value = sExpiryDate
        End If
        oCmd.Parameters.Add(paramExpiryDate)

        Dim paramReplenishmentDate As SqlParameter = New SqlParameter("@ReplenishmentDate", SqlDbType.SmallDateTime)
        Dim sReplenishmentDate As String
        sReplenishmentDate = tbReplenishmentDate.Text.Trim
        If sReplenishmentDate <> "" Then
            Try
                sReplenishmentDate = DateTime.Parse(sReplenishmentDate)
            Catch ex As Exception
                lblError.Text = "ERROR: Invalid Renewal / Review Date"
                Exit Sub
            End Try
        End If
        If sReplenishmentDate = "" Then
            paramReplenishmentDate.Value = Nothing
        Else
            paramReplenishmentDate.Value = sReplenishmentDate
        End If
        oCmd.Parameters.Add(paramReplenishmentDate)
      
        Dim paramSerialNumbers As SqlParameter = New SqlParameter("@SerialNumbersFlag", SqlDbType.NVarChar, 1)
        If chkProspectusNumbers.Checked Then
            paramSerialNumbers.Value = "Y"
        Else
            paramSerialNumbers.Value = "N"
        End If
        oCmd.Parameters.Add(paramSerialNumbers)
        Dim paramAdRotatorText As SqlParameter = New SqlParameter("@AdRotatorText", SqlDbType.NVarChar, 120)
        paramAdRotatorText.Value = txtAdRotatorText.Text
        oCmd.Parameters.Add(paramAdRotatorText)
        Dim paramWebsiteAdRotatorFlag As SqlParameter = New SqlParameter("@WebsiteAdRotatorFlag", SqlDbType.Bit)
        If chkAdRotator.Checked Then
            paramWebsiteAdRotatorFlag.Value = 1
        Else
            paramWebsiteAdRotatorFlag.Value = 0
        End If
        oCmd.Parameters.Add(paramWebsiteAdRotatorFlag)
        Dim paramNotes As SqlParameter = New SqlParameter("@Notes", SqlDbType.NVarChar, 1000)
        paramNotes.Value = txtNotes.Text
        oCmd.Parameters.Add(paramNotes)

        Dim paramViewOnWebForm As SqlParameter = New SqlParameter("@ViewOnWebForm", SqlDbType.Bit)
        If chkViewOnWebForm.Checked Then
            paramViewOnWebForm.Value = 1
        Else
            paramViewOnWebForm.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWebForm)
  
        Dim paramViewOnWebFormDE As SqlParameter = New SqlParameter("@Flag1", SqlDbType.Bit)
        If cbViewOnWebFormDE.Checked Then
            paramViewOnWebFormDE.Value = 1
        Else
            paramViewOnWebFormDE.Value = 0
        End If
        oCmd.Parameters.Add(paramViewOnWebFormDE)

        Dim paramFlag2 As SqlParameter = New SqlParameter("@Flag2", SqlDbType.Bit)
        paramFlag2.Value = 0
        oCmd.Parameters.Add(paramFlag2)

        Dim paramDefaultAccessFlag As SqlParameter = New SqlParameter("@DefaultAccessFlag", SqlDbType.Bit)
        paramDefaultAccessFlag.Value = Not gbExplicitProductPermissions
        oCmd.Parameters.Add(paramDefaultAccessFlag)

        Dim paramRotationProductKey As SqlParameter = New SqlParameter("@RotationProductKey", SqlDbType.Int, 4)
        paramRotationProductKey.Value = System.Data.SqlTypes.SqlInt32.Null
        oCmd.Parameters.Add(paramRotationProductKey)

        Dim paramInactivityAlertDays As SqlParameter = New SqlParameter("@InactivityAlertDays", SqlDbType.Int, 4)
        If IsNumeric(tbInactivityAlertDays.Text) Then
            paramInactivityAlertDays.Value = CLng(tbInactivityAlertDays.Text)
        Else
            paramInactivityAlertDays.Value = 0
        End If
        oCmd.Parameters.Add(paramInactivityAlertDays)
      
        Dim paramCalendarManaged As SqlParameter = New SqlParameter("@CalendarManaged", SqlDbType.Bit)
        If cbCalendarManaged.Checked Then
            paramCalendarManaged.Value = 1
        Else
            paramCalendarManaged.Value = 0
        End If
        oCmd.Parameters.Add(paramCalendarManaged)

        Dim paramOnDemand As SqlParameter = New SqlParameter("@OnDemand", SqlDbType.Int)
        paramOnDemand.Value = 0
        oCmd.Parameters.Add(paramOnDemand)
        
        Dim paramOnDemandPriceList As SqlParameter = New SqlParameter("@OnDemandPriceList", SqlDbType.Int)
        paramOnDemandPriceList.Value = 0
        oCmd.Parameters.Add(paramOnDemandPriceList)
        
        Dim paramCustomLetter As SqlParameter = New SqlParameter("@CustomLetter", SqlDbType.Bit)
        If pbCustomLetters Then
            If cbCustomLetter.Checked Then
                paramCustomLetter.Value = 1
            Else
                paramCustomLetter.Value = 0
            End If
        Else
            paramCustomLetter.Value = 0
        End If
        oCmd.Parameters.Add(paramCustomLetter)
        
        Dim paramProductCredits As SqlParameter = New SqlParameter("@ProductCredits", SqlDbType.Bit)
        If pbProductCredits Then
            If cbProductCredits.Checked Then
                paramProductCredits.Value = 1
            Else
                paramProductCredits.Value = 0
            End If
        Else
            paramProductCredits.Value = 0
        End If
        oCmd.Parameters.Add(paramProductCredits)

        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramProductKey)
  
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            plProductKey = CLng(oCmd.Parameters("@ProductKey").Value)
            Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
            Call ShowMainPanel()
        Catch ex As SqlException
            If ex.Number = 2627 Then
                lblError.Text = "ERROR: A record already exists with the same product CODE and DATE combination"
            Else
                lblError.Text = ex.ToString
            End If
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Function bSetExplicitProductPermissionsFlag() As Boolean
        Dim oDataTable As New DataTable()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_ExplicitProductPermissions_GetFlag", oConn)

        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            NotifyException("bSetExplicitProductPermissionsFlag", "Unable to retrieve ExplicitProductPermissions Flag because of an unexpected system error", ex, True, "Please retry - if the problem persists notify your Account Handler")
            bSetExplicitProductPermissionsFlag = False
        Finally
            oConn.Close()
            If IsDBNull(oDataTable.Rows(0).Item(0)) Then
                bSetExplicitProductPermissionsFlag = False
            Else
                bSetExplicitProductPermissionsFlag = oDataTable.Rows(0).Item(0)
            End If
        End Try
    End Function
  
    Protected Sub DeleteProduct()
        lblError.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_Delete", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
  
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)
        Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
        paramProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
            Call GetProductNumbers()
            Call ShowMainPanel()
        Catch ex As SqlException
            lblError.Text = "Error in DeleteProduct: " & ex.Message
        Finally
            oConn.Close()
            ResetForm()
        End Try
    End Sub
  
    Protected Sub GetProductNumbers()
        Dim oConn As New SqlConnection(gsConn)
        Dim sStoredProcedure As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sStoredProcedure = "spASPNET_Product_GetNumbersOwned"
        Else
            sStoredProcedure = "spASPNET_Product_GetNumbers"
        End If
        Dim oCmd As New SqlCommand(sStoredProcedure, oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim ParamCustomerKey As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
        ParamCustomerKey.Value = CLng(Session("CustomerKey"))

        If Session("UserType").ToString.ToLower.Contains("owner") Then
            Dim ParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int, 4)
            ParamUserKey.Value = Session("UserKey")
        End If

        Dim ParamProductCount As SqlParameter = oCmd.Parameters.Add("@ProductCount", SqlDbType.Int, 4)
        ParamProductCount.Direction = ParameterDirection.Output

        Dim ParamArchivedProductCount As SqlParameter = oCmd.Parameters.Add("@ArchivedProductCount", SqlDbType.Int, 4)
        ParamArchivedProductCount.Direction = ParameterDirection.Output

        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
            Dim nProductCount As Integer = CInt(ParamProductCount.Value)
            If nProductCount = 1 Then
                lblProductCountText.Text = "live product and"
            Else
                lblProductCountText.Text = "live products and"
            End If
            lblProductCount.Text = ParamProductCount.Value

            Dim nArchivedProductCount As Integer = CInt(ParamArchivedProductCount.Value)
            If nArchivedProductCount = 1 Then
                lblArchivedProductCountText.Text = "archived product"
            Else
                lblArchivedProductCountText.Text = "archived products"
            End If
            lblArchivedProductCount.Text = ParamArchivedProductCount.Value
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub BindProductUserProfileGrid(ByVal sSearchCriteria As String, ByVal sSortOrder As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        sSearchCriteria = sSearchCriteria.Trim
        Dim oAdapter As New SqlDataAdapter("", oConn)
        If sSearchCriteria.Length = 0 OrElse sSearchCriteria = "_" Then
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetUserProfilesFromKey"
        Else
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetUserProfilesMatchingSearch"
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
        End If
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = plProductKey
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        Try
            oAdapter.Fill(oDataTable)

            If oDataTable.Rows.Count > 0 Then
                lblProductProfileMessage.Text = ""
                Dim oDataView As DataView = oDataTable.DefaultView
                oDataView.Sort = sSortOrder
                grid_ProductUsers.DataSource = oDataView
                grid_ProductUsers.DataBind()
                grid_ProductUsers.Visible = True
                tblSaveCancelProductProfile.Visible = True
                lblLegendNoMatchingRecords.Visible = False
            Else
                grid_ProductUsers.Visible = False
                tblSaveCancelProductProfile.Visible = False
                lblProductProfileMessage.Text = "No users found"
                lblLegendNoMatchingRecords.Visible = True
            End If
        Catch ex As SqlException
            lblProductProfileMessage.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub SortProductUsersGrid(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        psSortValue = E.SortExpression
        grid_ProductUsers.CurrentPageIndex = 0
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
    End Sub
  
    Protected Sub btnToggleAllowToPickCheckboxes_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        For Each dgi In grid_ProductUsers.Items
            cb = dgi.FindControl("cbAllowToPick")
            If b.Text = "select all" Then
                cb.Checked = True
            Else
                cb.Checked = False
            End If
        Next
        If b.Text = "select all" Then
            b.Text() = "clear all"
        Else
            b.Text = "select all"
        End If
    End Sub
  
    Protected Sub btnToggleApplyMaxGrabCheckboxes_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        For Each dgi In grid_ProductUsers.Items
            cb = dgi.FindControl("cbApplyMaxGrab")
            If b.Text = "select all" Then
                cb.Checked = True
            Else
                cb.Checked = False
            End If
        Next
        If b.Text = "select all" Then
            b.Text() = "clear all"
        Else
            b.Text = "select all"
        End If
    End Sub

    Protected Sub btn_ApplyMaxGrabQty_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim tb As TextBox
        If IsNumeric(txtDefaultGrabQty.Text) AndAlso (CInt(txtDefaultGrabQty.Text) >= 0 And CInt(txtDefaultGrabQty.Text) <= 99999) Then
            For Each dgi In grid_ProductUsers.Items
                tb = dgi.FindControl("txtMaxGrabQty")
                tb.Text = txtDefaultGrabQty.Text
            Next
        End If
    End Sub

    Protected Sub btnSaveProductUserProfileChanges_click(ByVal s As Object, ByVal e As EventArgs)
        Call SaveProductUserProfileChanges()
    End Sub
  
    Protected Sub SaveProductUserProfileChanges()
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim tb As TextBox
        Dim bMaxGrabQtyAllValid As Boolean = True
        For Each dgi In grid_ProductUsers.Items
            tb = dgi.FindControl("txtMaxGrabQty")
            Dim sMaxGrabQty As String = tb.Text.Trim
            If sMaxGrabQty.Length > 0 AndAlso Not IsNumeric(sMaxGrabQty) Then
                tb.ForeColor = Drawing.Color.Red
                tb.Font.Bold = True
                bMaxGrabQtyAllValid = False
            Else
                tb.ForeColor = Nothing
                tb.Font.Bold = Nothing
            End If
        Next
        If Not bMaxGrabQtyAllValid Then
            WebMsgBox.Show("One or more max grab quantities has a value that is not a valid number - check max grab values are blank or in the range 0 - 99999")
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetUserProductProfile", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            oCmd.Parameters.Add(paramUserKey)
            Dim paramProductKey As SqlParameter = New SqlParameter("@ProductKey", SqlDbType.Int, 4)
            paramProductKey.Value = plProductKey
            oCmd.Parameters.Add(paramProductKey)
            Dim paramAbleToView As SqlParameter = New SqlParameter("@AbleToView", SqlDbType.Bit)
            paramAbleToView.Value = 1
            oCmd.Parameters.Add(paramAbleToView)
            Dim paramAbleToPick As SqlParameter = New SqlParameter("@AbleToPick", SqlDbType.Bit)
            oCmd.Parameters.Add(paramAbleToPick)
            Dim paramAbleToEdit As SqlParameter = New SqlParameter("@AbleToEdit", SqlDbType.Bit)
            paramAbleToEdit.Value = 0
            oCmd.Parameters.Add(paramAbleToEdit)
            Dim paramAbleToArchive As SqlParameter = New SqlParameter("@AbleToArchive", SqlDbType.Bit)
            paramAbleToArchive.Value = 0
            oCmd.Parameters.Add(paramAbleToArchive)
            Dim paramAbleToDelete As SqlParameter = New SqlParameter("@AbleToDelete", SqlDbType.Bit)
            paramAbleToDelete.Value = 0
            oCmd.Parameters.Add(paramAbleToDelete)
            Dim paramApplyMaxGrab As SqlParameter = New SqlParameter("@ApplyMaxGrab", SqlDbType.Bit)
            oCmd.Parameters.Add(paramApplyMaxGrab)
            Dim paramMaxGrabQty As SqlParameter = New SqlParameter("@MaxGrabQty", SqlDbType.Int, 4)
            oCmd.Parameters.Add(paramMaxGrabQty)
            Try
                oConn.Open()
                oCmd.Connection = oConn
                For Each dgi In grid_ProductUsers.Items
                    oCmd.Parameters("@UserKey").Value = dgi.Cells(0).Text
                    cb = dgi.FindControl("cbAllowToPick")
                    oCmd.Parameters("@AbleToPick").Value = cb.Checked
                    cb = dgi.FindControl("cbApplyMaxGrab")
                    oCmd.Parameters("@ApplyMaxGrab").Value = cb.Checked
                    tb = dgi.FindControl("txtMaxGrabQty")
                    Dim sTest = tb.Text.Trim
                    If tb.Text.Trim.Length > 0 Then
                        oCmd.Parameters("@MaxGrabQty").Value = CInt(tb.Text)
                    Else
                        oCmd.Parameters("@MaxGrabQty").Value = 0
                    End If
                    oCmd.ExecuteNonQuery()
                Next
            Catch ex As Exception
                NotifyException("SaveProductUserProfileChanges", "Unable to update user permissions because of an unexpected system error", ex, True, "Please retry - if the problem persists notify your Account Handler")
            Finally
                oConn.Close()
            End Try
            Call ReturnToProductDetail()
        End If
    End Sub
  
    Protected Sub ResetCategoryDropdowns()
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        pbIsAddingSubSubCategory = False
        Call CheckVisibility()
    End Sub
  
    Protected Sub ddlCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_CATEGORY Then
            pbIsAddingCategory = True
        End If
        Call CheckVisibility()
    End Sub

    Protected Sub ddlSubCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_SUBCATEGORY Then
            pbIsAddingSubCategory = True
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub ddlSubSubCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedItem.Text = NEW_SUBCATEGORY Then
            pbIsAddingSubSubCategory = True
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub CheckVisibility()
        pbIsAddingCategory = pbIsAddingCategory
        pbIsAddingSubCategory = pbIsAddingSubCategory
        pbIsAddingSubSubCategory = pbIsAddingSubSubCategory
    End Sub

    Protected Sub ddlCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlCategory()
        'If pbIsAddingNew AndAlso IsKODDFIS() Then
        '    Dim sCategory = ddlKODDFISCategory.SelectedItem.Text
        '    For i As Integer = 2 To ddlCategory.Items.Count - 1
        '        If ddlCategory.Items(i).Text = sCategory Then
        '            ddlCategory.SelectedIndex = i
        '            Exit For
        '        End If
        '    Next
        'End If
    End Sub

    Protected Sub ddlSubCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlSubCategory()
    End Sub

    Protected Sub ddlSubSubCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        gbDataBound = True
        Call AdjustddlSubSubCategory()
    End Sub

    Protected Function CanCreateNewCategories() As Boolean
        If pbProductOwners AndAlso Session("UserType").ToString.ToLower <> "superuser" Then
            CanCreateNewCategories = False
        Else
            CanCreateNewCategories = True
        End If
    End Function
  
    Protected Sub AdjustddlCategory()
        If CanCreateNewCategories() Then
            ddlCategory.Items.Insert(0, NEW_CATEGORY)
        End If
        'ddlCategory.Items.Insert(0, NEW_CATEGORY)
        ddlCategory.Items.Insert(0, "")
        If Not pbIsAddingCategory Then
            If pbIsAddingNew Then
                ddlCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlCategory.Items.Count - 1
                    'If ddlCategory.Items(i).Text = hidCategory.Value.ToString Then
                    If ddlCategory.Items(i).Text.Trim.ToLower = hidCategory.Value.ToString.Trim.ToLower Then   'CN 3FEB09
                        ddlCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Protected Function ContainsNewSubCategory(ByVal ddl As DropDownList) As Boolean
        ContainsNewSubCategory = False
        If ddl.Items.Count > 0 Then
            For Each li As ListItem In ddl.Items
                If li.Text = NEW_SUBCATEGORY Then
                    ContainsNewSubCategory = True
                    Exit Function
                End If
            Next
        End If
    End Function
  
    Protected Sub AdjustddlSubCategory()
        'If ddlSubCategory.Items.Count = 0 OrElse ddlSubCategory.Items(1).Text <> NEW_SUBCATEGORY Then
        If Not ContainsNewSubCategory(ddlSubCategory) Then
            If CanCreateNewCategories() Then
                ddlSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            End If
            'ddlSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            ddlSubCategory.Items.Insert(0, "")
        End If
        If Not pbIsAddingSubCategory Then
            If pbIsAddingNew Then
                ddlSubCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlSubCategory.Items.Count - 1
                    'If ddlSubCategory.Items(i).Text = hidSubCategory.Value.ToString Then
                    If ddlSubCategory.Items(i).Text.Trim.ToLower = hidSubCategory.Value.ToString.Trim.ToLower Then   'CN 3FEB09
                        ddlSubCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
  
    Protected Sub AdjustddlSubSubCategory()
        If Not ContainsNewSubCategory(ddlSubSubCategory) Then
            If CanCreateNewCategories() Then
                ddlSubSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            End If
            'ddlSubSubCategory.Items.Insert(0, NEW_SUBCATEGORY)
            ddlSubSubCategory.Items.Insert(0, "")
        End If
        If Not pbIsAddingSubSubCategory Then
            If pbIsAddingNew Then
                ddlSubSubCategory.SelectedIndex = 0
            Else
                Dim i As Integer
                For i = 0 To ddlSubSubCategory.Items.Count - 1
                    'If ddlSubSubCategory.Items(i).Text = hidSubSubCategory.Value.ToString Then
                    If ddlSubSubCategory.Items(i).Text.Trim.ToLower = hidSubSubCategory.Value.ToString.Trim.ToLower Then   'CN 3FEB09
                        ddlSubSubCategory.SelectedIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
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
  
    Protected Sub ResetForm()
        txtProductCode.Text = String.Empty
        txtProductDate.Text = String.Empty
        txtMinStockLevel.Text = String.Empty
        txtDescription.Text = String.Empty
        txtItemsPerBox.Text = String.Empty
        txtCategory.Text = String.Empty
        txtSubCategory.Text = String.Empty
        tbSubSubCategory.Text = String.Empty
        txtUnitValue.Text = String.Empty
        txtLanguage.Text = String.Empty
        txtUnitWeight.Text = String.Empty
        txtMisc1.Text = String.Empty
        txtMisc2.Text = String.Empty
        chkArchivedFlag.Checked = False
        tbExpiryDate.Text = String.Empty
        tbReplenishmentDate.Text = String.Empty
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        chkProspectusNumbers.Checked = False
        txtAdRotatorText.Text = String.Empty
        chkAdRotator.Checked = False
        chkViewOnWebForm.Checked = False
        cbRequiresAuth.Checked = False
        cbCalendarManaged.Checked = False
        lnkbtnPreAuthorise.Visible = False
        txtNotes.Text = String.Empty
        hlnk_PDF.ImageUrl = String.Empty
        hlnk_PDF.NavigateUrl = String.Empty
        hlnk_DetailThumb.ImageUrl = String.Empty
        hlnk_DetailThumb.NavigateUrl = String.Empty
        tbSellingPrice.Text = String.Empty
        txtDepartment.Text = String.Empty
        ddlProductGroup.SelectedIndex = 0
        cbOnDemand.Checked = False
        If IsJupiter() Then
            'ddlPrintType.SelectedIndex = 0
            ddlPODPageCount.SelectedIndex = 0
        End If
    End Sub

    Protected Sub btnPendingAuthorisations_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If pbOrderAuthorisation Then
            gvAuthoriseOrder.PageIndex = 0
            Call GetPendingOrderAuthorisations()
            If gvAuthoriseOrder.Rows.Count > 0 Then
                Call ShowAuthoriseOrderPanel()
            Else
                Call ShowMainPanel()
                WebMsgBox.Show("You have no authorisation requests awaiting approval.")
            End If
        Else
            Call GetPendingProductAuthorisations()
            If gvAuthoriseProduct.Rows.Count > 0 Then
                Call ShowAuthoriseProductPanel()
            Else
                Call ShowMainPanel()
                WebMsgBox.Show("You have no authorisation requests awaiting approval.")
            End If
        End If
    End Sub

    Protected Sub GetPendingProductAuthorisations()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetPendingAuthorisationRequests", oConn)
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@AuthoriserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@AuthoriserKey").Value = Session("UserKey")

        oAdapter.Fill(oDatatable)
        gvAuthoriseProduct.DataSource = oDatatable
        gvAuthoriseProduct.DataBind()
        If gvAuthoriseProduct.Rows.Count > 0 Then
            Dim gvr As GridViewRow = gvAuthoriseProduct.Rows(0)
            Dim tb As TextBox
            tb = gvr.FindControl("tbQuantityAuthorised")
            tb.Focus()
        End If
    End Sub
  
    Protected Sub btnAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DoAuthorisations()
    End Sub
  
    Protected Sub DoAuthorisations()
        Dim oConn As New SqlConnection(gsConn)
        Dim cbAuthoriseProduct As CheckBox
        Dim tbQuantityAuthorised As TextBox
        Dim nAuthorisationCount As Integer = 0
        Dim hidAuthoriseId As HiddenField
        Dim bIsValid As Boolean = True
        lblAuthErrorMessage.Text = ""
        For Each gvr As GridViewRow In gvAuthoriseProduct.Rows
            cbAuthoriseProduct = gvr.FindControl("cbAuthoriseProduct")
            If cbAuthoriseProduct.Checked Then
                tbQuantityAuthorised = gvr.FindControl("tbQuantityAuthorised")
                If Not IsNumeric(tbQuantityAuthorised.Text) Then
                    bIsValid = False
                    lblAuthErrorMessage.Text = "Please check quantities"
                End If
                Dim tbDuration As TextBox = gvr.FindControl("tbDuration")
                tbDuration.Text = tbDuration.Text.Trim
                If Not (IsNumeric(tbDuration.Text) Or tbDuration.Text = "unlimited" Or tbDuration.Text = String.Empty) Then
                    bIsValid = False
                    lblAuthErrorMessage.Text = "Please check durations"
                End If
            End If
        Next
        If bIsValid Then
            For Each gvr As GridViewRow In gvAuthoriseProduct.Rows
                cbAuthoriseProduct = gvr.FindControl("cbAuthoriseProduct")
                tbQuantityAuthorised = gvr.FindControl("tbQuantityAuthorised")
                hidAuthoriseId = gvr.FindControl("hidAuthoriseId")
                If cbAuthoriseProduct.Checked Or rblAuthoriseAction.SelectedValue = "decline" Then
                    Dim bResult As Boolean = cbAuthoriseProduct.Checked
                    Dim tbDuration As TextBox = gvr.FindControl("tbDuration")
                    tbDuration.Text = tbDuration.Text.Trim
                    Dim dtExpiryDate As DateTime
                    If IsNumeric(tbDuration.Text) Then
                        Dim tsDuration As TimeSpan = TimeSpan.FromDays(CInt(tbDuration.Text))
                        dtExpiryDate = Now() + tsDuration
                    Else
                        Dim tsDuration As TimeSpan = TimeSpan.FromDays(5000)
                        dtExpiryDate = Now() + tsDuration
                    End If
                    Dim sExpiryDate As String = dtExpiryDate.ToString
              
                    Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetAuthorisation", oConn)
                    oCmd.CommandType = CommandType.StoredProcedure

                    Dim paramAuthorisationKey As SqlParameter = New SqlParameter("@AuthorisationKey", SqlDbType.Int)
                    paramAuthorisationKey.Value = hidAuthoriseId.Value
                    oCmd.Parameters.Add(paramAuthorisationKey)
              
                    Dim paramResult As SqlParameter = New SqlParameter("@Result", SqlDbType.Bit)
                    paramResult.Value = bResult
                    oCmd.Parameters.Add(paramResult)
              
                    Dim paramQuantityAuthorised As SqlParameter = New SqlParameter("@Quantity", SqlDbType.Int)
                    paramQuantityAuthorised.Value = tbQuantityAuthorised.Text
                    oCmd.Parameters.Add(paramQuantityAuthorised)
              
                    Dim paramExpiryDate As SqlParameter = New SqlParameter("@Expiry", SqlDbType.SmallDateTime)
                    paramExpiryDate.Value = sExpiryDate
                    oCmd.Parameters.Add(paramExpiryDate)

                    Dim paramMessage As SqlParameter = New SqlParameter("@Message", SqlDbType.VarChar, 4000)
                    paramMessage.Value = System.Data.SqlTypes.SqlString.Null
                    oCmd.Parameters.Add(paramMessage)

                    Try
                        oConn.Open()
                        oCmd.Connection = oConn
                        oCmd.ExecuteNonQuery()
                        nAuthorisationCount += 1
                    Catch ex As SqlException
                        WebMsgBox.Show("Unable to set authorisation status - aborting")
                    Finally
                        oConn.Close()
                    End Try
                End If
            Next
        End If
        If bIsValid Then
            Dim sPlural As String
            If nAuthorisationCount Then
                sPlural = ""
            Else
                sPlural = "s"
            End If
            WebMsgBox.Show(nAuthorisationCount.ToString & " authorisation" & sPlural & " completed")
            Call ShowMainPanel()
        End If
    End Sub
  
    Protected Sub btnSelectAuthoriseProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cbAuthoriseProduct As CheckBox
        Dim b As Button = sender
        Dim bValue As Boolean

        If b.Text = "select all" Then
            bValue = True
            b.Text = "clear all"
        Else
            bValue = False
            b.Text = "select all"
        End If

        For Each gvr As GridViewRow In gvAuthoriseProduct.Rows
            cbAuthoriseProduct = DirectCast(gvr.Controls(11).Controls(1), CheckBox)
            'cbAuthorise = DirectCast(FindControl("cbAuthoriseProduct"), CheckBox)
            cbAuthoriseProduct.Checked = bValue
        Next
    End Sub
  
    Protected Sub cbRequiresAuth_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        cbChangeAuthoriserOnAll.Checked = False
        cbChangeAuthoriserOnAllSelective.Checked = False
        If cb.Checked Then
            If cbCalendarManaged.Checked Then
                WebMsgBox.Show("Cannot set the Requires Authorisation attribute on a Calendar Managed product. Only one of these facilities can be used. Remove the Calendar Managed attribute if you want to make this product authorisable.")
                cb.Checked = False
                Exit Sub
            End If
            Call PopulateSuperUserDropdown(ddlAssignAuthoriser)
            Call ShowMakeAuthorisable()
            lnkbtnPreAuthorise.Visible = True
        Else
            Dim sCurrentAuthoriser As String = GetAuthoriser()
            lblCurrentAuthoriser.Text = sCurrentAuthoriser
            cbChangeAuthoriserOnAllSelective.Text = "change on all <i>req auth</i> products where " & sCurrentAuthoriser & " <b>is the current authoriser</b>"
            btnSaveNewAuthoriser.Enabled = False
            Call PopulateSuperUserDropdown(ddlModifyAuthoriser)
            For i As Integer = 1 To ddlModifyAuthoriser.Items.Count - 1
                If ddlModifyAuthoriser.Items(i).Text.StartsWith(sCurrentAuthoriser) Then
                    ddlModifyAuthoriser.Items.RemoveAt(i)
                    Exit For
                End If
            Next
            Call ShowRemoveAuthorisable()
            lnkbtnPreAuthorise.Visible = False
        End If
        Call CheckVisibility()
    End Sub
  
    Protected Sub PopulateSuperUserDropdown(ByVal ddl As DropDownList)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetAllSuperUsersForCustomer2", oConn)
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.Fill(oDatatable)
        ddl.Items.Clear()
        Dim li As New ListItem
        li.Text = "-- select authoriser --"
        li.Value = 0
        ddl.Items.Add(li)

        For Each dr As DataRow In oDatatable.Rows
            Dim li2 As New ListItem
            Dim sFirstName As String = dr("FirstName")
            Dim sLastName As String = dr("LastName")
            li2.Text = Char.ToUpper(sFirstName(0)) & sFirstName.Substring(1) & " " & Char.ToUpper(sLastName(0)) & sLastName.Substring(1) & "  (" & dr("UserId") & ")"
            li2.Value = dr("key")
            ddl.Items.Add(li2)
        Next
    End Sub

    Protected Sub btnSaveAuthorisable_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlAssignAuthoriser.SelectedValue > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetAuthorisable", oConn)
            oCmd.CommandType = CommandType.StoredProcedure

            Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
            paramLogisticProductKey.Value = plProductKey
            oCmd.Parameters.Add(paramLogisticProductKey)
      
            Dim paramDefaultAuthorisationGrantTimeoutHours As SqlParameter = New SqlParameter("@DefaultAuthorisationGrantTimeoutHours", SqlDbType.Int)
            paramDefaultAuthorisationGrantTimeoutHours.Value = 0
            oCmd.Parameters.Add(paramDefaultAuthorisationGrantTimeoutHours)
      
            Dim paramDefaultAuthorisationLifetimeHours As SqlParameter = New SqlParameter("@DefaultAuthorisationLifetimeHours", SqlDbType.Int)
            paramDefaultAuthorisationLifetimeHours.Value = 0
            oCmd.Parameters.Add(paramDefaultAuthorisationLifetimeHours)
      
            Dim paramAuthoriser As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int)
            paramAuthoriser.Value = ddlAssignAuthoriser.SelectedValue
            oCmd.Parameters.Add(paramAuthoriser)
      
            Try
                oConn.Open()
                oCmd.Connection = oConn
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                WebMsgBox.Show("Unable to set product authorisable - aborting.")
            Finally
                oConn.Close()
            End Try
            Call ShowProductDetail()
        Else
            WebMsgBox.Show("Please select an authoriser for this product.")
        End If
    End Sub
  
    Protected Sub btnRequestAuthGoBack_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbRequiresAuth.Checked = False
        lnkbtnPreAuthorise.Visible = False
        Call ShowProductDetail()
    End Sub
  
    Protected Sub btnRemoveRequestAuthGoBack_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbRequiresAuth.Checked = True
        lnkbtnPreAuthorise.Visible = True
        Call ShowProductDetail()
    End Sub
  
    Protected Sub btnRemoveAuthorisable_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_RemoveAuthorisable", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
      
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Unable to remove product authorisable attribute - aborting.")
        Finally
            oConn.Close()
        End Try
        Call ShowProductDetail()
    End Sub
  
    Protected Sub BindPreAuthoriseGrid(ByVal sSearchCriteria As String, ByVal sSortOrder As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        sSearchCriteria = sSearchCriteria.Trim
        Dim oAdapter As New SqlDataAdapter("", oConn)
        If sSearchCriteria.Length = 0 OrElse sSearchCriteria = "_" Then
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetAllPreAuthorisationUsers2"
        Else
            oAdapter.SelectCommand.CommandText = "spASPNET_Product_GetPreAuthorisationUsersMatchingSearch2"
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
        End If
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = plProductKey
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        Try
            oAdapter.Fill(oDataTable)

            If oDataTable.Rows.Count > 0 Then
                lblProductProfileMessage.Text = ""
                Dim oDataView As DataView = oDataTable.DefaultView
                oDataView.Sort = sSortOrder
                dgPreAuthorise.DataSource = oDataView
                dgPreAuthorise.DataBind()
                dgPreAuthorise.Visible = True
            Else
                dgPreAuthorise.Visible = False
                lblPreAuthoriseMessage.Text = "No users found"
            End If
        Catch ex As SqlException
            lblPreAuthoriseMessage.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub SortPreAuthoriseGrid(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Call BindPreAuthoriseGrid(tbPreAuthoriseUserSearch.Text, E.SortExpression)
    End Sub
  
    Protected Sub btnToggleAuthoriseCheckboxes_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        Dim b As Button = s
        For Each dgi In dgPreAuthorise.Items
            cb = dgi.FindControl("cbPreAuthorise")
            If b.Text = "select all" Then
                cb.Checked = True
            Else
                cb.Checked = False
            End If
        Next
        If b.Text = "select all" Then
            b.Text() = "clear all"
        Else
            b.Text = "select all"
        End If
    End Sub
  
    Protected Sub btnApplyPreAuthoriseQty_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim tb As TextBox
        If IsNumeric(tbDefaultPreAuthoriseQty.Text) AndAlso (CInt(tbDefaultPreAuthoriseQty.Text) >= 0 And CInt(tbDefaultPreAuthoriseQty.Text) <= 99999) Then
            For Each dgi In dgPreAuthorise.Items
                tb = dgi.FindControl("tbPreAuthoriseQty")
                tb.Text = tbDefaultPreAuthoriseQty.Text
            Next
        End If
    End Sub

    Protected Sub btnApplyDuration_Click(ByVal s As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim tb As TextBox
        If IsNumeric(tbDefaultDuration.Text) AndAlso (CInt(tbDefaultDuration.Text) >= 0 And CInt(tbDefaultDuration.Text) <= 99999) Then
            For Each dgi In dgPreAuthorise.Items
                tb = dgi.FindControl("tbDuration")
                tb.Text = tbDefaultDuration.Text
            Next
        End If
    End Sub

    Protected Sub btnSavePreAuthoriseChanges_Click(ByVal s As Object, ByVal e As EventArgs)
        Call SavePreAuthoriseChanges()
    End Sub
  
    Protected Sub SavePreAuthoriseChanges()
        Dim dgi As DataGridItem
      
        Dim cbPreAuthorise As CheckBox
        Dim tbPreAuthoriseQty As TextBox
        Dim tbDuration As TextBox
      
        Dim bPreAuthoriseFieldsAllValid As Boolean = True
        Dim nPreAuthoriseCount As Integer = 0
      
        For Each dgi In dgPreAuthorise.Items
            cbPreAuthorise = dgi.FindControl("cbPreAuthorise")
            If cbPreAuthorise.Checked Then
              
                tbPreAuthoriseQty = dgi.FindControl("tbPreAuthoriseQty")
                Dim sPreAuthoriseQty As String = tbPreAuthoriseQty.Text.Trim
                If sPreAuthoriseQty.Length = 0 OrElse Not IsNumeric(sPreAuthoriseQty) Then
                    tbPreAuthoriseQty.BackColor = Drawing.Color.Yellow
                    tbPreAuthoriseQty.ForeColor = Drawing.Color.Red
                    tbPreAuthoriseQty.Font.Bold = True
                    bPreAuthoriseFieldsAllValid = False
                Else
                    tbPreAuthoriseQty.ForeColor = Nothing
                    tbPreAuthoriseQty.Font.Bold = Nothing
                End If
              
                tbDuration = dgi.FindControl("tbDuration")
                Dim sDuration As String = tbDuration.Text.Trim
                If sDuration.Length = 0 AndAlso Not IsNumeric(sDuration) Then
                    tbDuration.BackColor = Drawing.Color.Yellow
                    tbDuration.ForeColor = Drawing.Color.Red
                    tbDuration.Font.Bold = True
                    bPreAuthoriseFieldsAllValid = False
                Else
                    tbDuration.ForeColor = Nothing
                    tbDuration.Font.Bold = Nothing
                End If
            End If
        Next

        If Not bPreAuthoriseFieldsAllValid Then
            WebMsgBox.Show("One or more authorise quantities has a value that is not a valid number - check authorise quantities are blank or in the range 0 - 99999")
        Else
            For Each dgi In dgPreAuthorise.Items
                Dim oConn As New SqlConnection(gsConn)
                cbPreAuthorise = dgi.FindControl("cbPreAuthorise")
                If cbPreAuthorise.Checked Then
                    tbPreAuthoriseQty = dgi.FindControl("tbPreAuthoriseQty")
                    tbDuration = dgi.FindControl("tbDuration")
                    Dim dtExpiryDate As DateTime
                    If IsNumeric(tbDuration.Text) Then
                        Dim tsDuration As TimeSpan = TimeSpan.FromDays(CInt(tbDuration.Text))
                        dtExpiryDate = Now() + tsDuration
                    Else
                        Dim tsDuration As TimeSpan = TimeSpan.FromDays(5000)
                        dtExpiryDate = Now() + tsDuration
                    End If
                    Dim sExpiryDate As String = dtExpiryDate.ToString
              
                    Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_SetPreAuthorisation", oConn)
                    oCmd.CommandType = CommandType.StoredProcedure

                    Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
                    paramLogisticProductKey.Value = plProductKey
                    oCmd.Parameters.Add(paramLogisticProductKey)
              
                    Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int)
                    paramUserProfileKey.Value = dgi.Cells(0).Text
                    oCmd.Parameters.Add(paramUserProfileKey)
              
                    Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
                    paramCustomerKey.Value = Session("CustomerKey")
                    oCmd.Parameters.Add(paramCustomerKey)
              
                    Dim paramQuantityAuthorised As SqlParameter = New SqlParameter("@Quantity", SqlDbType.Int)
                    paramQuantityAuthorised.Value = tbPreAuthoriseQty.Text
                    oCmd.Parameters.Add(paramQuantityAuthorised)
              
                    Dim paramExpiryDate As SqlParameter = New SqlParameter("@Expiry", SqlDbType.SmallDateTime)
                    paramExpiryDate.Value = sExpiryDate
                    oCmd.Parameters.Add(paramExpiryDate)

                    Try
                        oConn.Open()
                        oCmd.Connection = oConn
                        oCmd.ExecuteNonQuery()
                        nPreAuthoriseCount += 1
                    Catch ex As SqlException
                        WebMsgBox.Show("Unable to set authorisation status - aborting: " & ex.Message)
                    Finally
                        oConn.Close()
                    End Try
                End If
            Next
            If bPreAuthoriseFieldsAllValid Then
                If nPreAuthoriseCount > 0 Then
                    Dim sProductLegend As String = "product"
                    If nPreAuthoriseCount > 1 Then
                        sProductLegend += "s"
                    End If
                    WebMsgBox.Show(nPreAuthoriseCount.ToString & " " & sProductLegend & " modified")
                End If
                Call ReturnToProductDetail()
            End If
        End If
    End Sub
  
    Protected Sub lnkbtnPreAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPreAuthoriseUserSearch.Text = ""
        Call BindPreAuthoriseGrid(tbPreAuthoriseUserSearch.Text, "UserID")
        Call ShowPreAuthorise()
    End Sub
  
    Protected Sub btnPreAuthoriseShowAllUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPreAuthoriseUserSearch.Text = ""
        Call BindPreAuthoriseGrid(tbPreAuthoriseUserSearch.Text, "UserID")
    End Sub
  
    Protected Sub btnPreAuthoriseSearchUsers_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindPreAuthoriseGrid(tbPreAuthoriseUserSearch.Text, "UserID")
    End Sub
  
    Protected Sub btnProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroupsPanel()
    End Sub
  
    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
      
        ddlPrimaryProductGroupOwner.SelectedIndex = 0
        ddlDeputyProductGroupOwner.SelectedIndex = 0
      
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT * FROM ProductGroup WHERE ProductGroupName = '" & ddl.SelectedItem.Text.Replace("'", "''") & "' AND CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            oDataReader.Read()
            Dim nProductOwner1 As Integer
            Dim nProductOwner2 As Integer
            If IsDBNull(oDataReader("ProductOwner1")) Then
                nProductOwner1 = 0
            Else
                nProductOwner1 = oDataReader("ProductOwner1")
            End If
            If IsDBNull(oDataReader("ProductOwner2")) Then
                nProductOwner2 = 0
            Else
                nProductOwner2 = oDataReader("ProductOwner2")
            End If
            For i As Integer = 1 To ddlPrimaryProductGroupOwner.Items.Count - 1
                If ddlPrimaryProductGroupOwner.Items(i).Value = nProductOwner1 Then
                    ddlPrimaryProductGroupOwner.SelectedIndex = i
                    Exit For
                End If
            Next
            For i As Integer = 1 To ddlDeputyProductGroupOwner.Items.Count - 1
                If ddlDeputyProductGroupOwner.Items(i).Value = nProductOwner2 Then
                    ddlDeputyProductGroupOwner.SelectedIndex = i
                    Exit For
                End If
            Next
        Else
            WebMsgBox.Show("Internal error retrieving product group record")
        End If
        If ddl.SelectedIndex > 0 Then
            btnShowProductsInGroup.Visible = True
            gvProductsInGroup.Visible = False
        Else
            btnShowProductsInGroup.Visible = False
            gvProductsInGroup.Visible = False
        End If
    End Sub
  
    Protected Sub btnNewProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNewProductGroup.Text = String.Empty
        tbNewProductGroup.Focus()
        Call ShowNewProductGroupPanel()
    End Sub
  
    Protected Sub btnCreateNewProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbNewProductGroup.Text = tbNewProductGroup.Text.Trim
        If tbNewProductGroup.Text = String.Empty Then
            WebMsgBox.Show("Please enter a name for the group.")
            Exit Sub
        End If
        tbNewProductGroup.Text = tbNewProductGroup.Text.Trim
        If IsExistingProductGroupName(tbNewProductGroup.Text) Then
            WebMsgBox.Show("A product group with this name already exists.  Please choose an alternative name.")
        Else
            Call CreateNewProductGroup(tbNewProductGroup.Text)
            ddlProductGroup.Items.Clear()
            ddlAssignedProductGroup.Items.Clear()
            Call InitProductGroupControls()
            Call ShowProductGroupsPanel()
        End If
    End Sub

    Protected Sub btnRenameThisProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbRenamedProductGroup.Text = tbRenamedProductGroup.Text.Trim
        If IsExistingProductGroupName(tbRenamedProductGroup.Text) Then
            WebMsgBox.Show("A product group with this name already exists.  Please choose an alternative name.")
        Else
            Call RenameProductGroup(tbRenamedProductGroup.Text)
            ddlProductGroup.Items.Clear()
            ddlAssignedProductGroup.Items.Clear()
            Call InitProductGroupControls()
            Call ShowProductGroupsPanel()
        End If
    End Sub

    Protected Function sTimeStamp() As String
        sTimeStamp = Now
    End Function
  
    Protected Sub CreateNewProductGroup2(ByVal sProductGroupName As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "INSERT INTO ProductGroup (ProductGroupName, CustomerKey, ProductOwner1, ProductOwner2, CreatedOn, CreatedBy) VALUES ('"
        sSQL += sProductGroupName + "', " & Session("CustomerKey") & ", '', '', '" & sTimeStamp() & "', " & Session("UserKey") & ")"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("CreateNewProductGroup2 " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub CreateNewProductGroup(ByVal sProductGroupName As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_CreateProductGroup", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramProductGroupName As SqlParameter = New SqlParameter("@ProductGroupName", SqlDbType.NVarChar, 50)
        paramProductGroupName.Value = sProductGroupName
        oCmd.Parameters.Add(paramProductGroupName)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in CreateNewProductGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub RenameProductGroup2(ByVal sNewProductGroupName As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Dim sSQL As String = "UPDATE ProductGroup SET ProductGroupName = '" & sNewProductGroupName & "', LastUpdated = '" & sTimeStamp() & "' WHERE ProductGroupName = '" & lblProductGroupToRename.Text.Replace("'", "''") & "'"
        Try
            oConn.Open()
            oCmd = New SqlCommand(sSQL, oConn)
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in CreateNewProductGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub RenameProductGroup(ByVal sNewProductGroupName As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_RenameProductGroup", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCurrentProductGroupName As SqlParameter = New SqlParameter("@CurrentProductGroupName", SqlDbType.NVarChar, 50)
        paramCurrentProductGroupName.Value = lblProductGroupToRename.Text
        oCmd.Parameters.Add(paramCurrentProductGroupName)

        Dim paramNewProductGroupName As SqlParameter = New SqlParameter("@NewProductGroupName", SqlDbType.NVarChar, 50)
        paramNewProductGroupName.Value = sNewProductGroupName
        oCmd.Parameters.Add(paramNewProductGroupName)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in RenameProductGroup " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Function IsExistingProductGroupName(ByVal sProductGroupName As String) As Boolean
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String

        IsExistingProductGroupName = False

        sSQL = "SELECT * FROM ProductGroup WHERE ProductGroupName = '" & sProductGroupName.Replace("'", "''") & "' AND CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)

        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            IsExistingProductGroupName = True
        End If
        oConn.Close()
    End Function
  
    Protected Function GetProductOwner(ByVal sUserKey As String) As String
        Dim sFullName As String = String.Empty
        If sUserKey = "0" Then
            sFullName = "unassigned"
        Else
            Dim oDataReader As SqlDataReader = Nothing
            Dim oConn As New SqlConnection(gsConn)
            Dim sSQL As String = "SELECT * FROM UserProfile WHERE [Key] = " & sUserKey
            Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                sFullName = oDataReader("FirstName") & " " & oDataReader("LastName")
            Else
                sFullName = "not found"
            End If
            oConn.Close()
        End If
        GetProductOwner = sFullName
    End Function
  
    Protected Function GetProductOwners() As Dictionary(Of String, Integer)
        Dim dicProductOwners As New Dictionary(Of String, Integer)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        'If Not cbUnassignedProductOwnersOnly.Checked Then
        '    sSQL = "SELECT [Key], FirstName, LastName, UserId FROM UserProfile WHERE (Type LIKE '%owner%' OR Type LIKE '%superuser%') AND Status = 'Active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY LastName"
        'Else
        '    sSQL = "SELECT [Key], FirstName, LastName, UserId FROM UserProfile WHERE (Type LIKE '%owner%' OR Type LIKE '%superuser%') AND Status = 'Active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " AND NOT ([Key] IN (SELECT ProductOwner1 FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & ") OR [Key] IN (SELECT ProductOwner2 FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & ")) ORDER BY LastName"
        'End If
        If Not cbUnassignedProductOwnersOnly.Checked Then
            sSQL = "SELECT [Key], FirstName, LastName, UserId FROM UserProfile WHERE Type LIKE '%owner%' AND Status = 'Active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " ORDER BY LastName"
        Else
            sSQL = "SELECT [Key], FirstName, LastName, UserId FROM UserProfile WHERE Type LIKE '%owner%' AND Status = 'Active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " AND NOT ([Key] IN (SELECT ProductOwner1 FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & ") OR [Key] IN (SELECT ProductOwner2 FROM ProductGroup WHERE CustomerKey = " & Session("CustomerKey") & ")) ORDER BY LastName"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        While oDataReader.Read
            Dim sFullName As String = oDataReader("FirstName") & " " & oDataReader("LastName") & " (" & oDataReader("UserId") & ")"
            dicProductOwners.Add(sFullName, oDataReader("Key"))
        End While
        oConn.Close()
        GetProductOwners = dicProductOwners
    End Function
  
    Protected Sub btnBackFromProductGroupsToList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BackToProductListPanel()
    End Sub
  
    Protected Sub btnBackFromNewProductGroupToProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroupsPanel()
    End Sub

    Protected Sub btnBackFromRenameProductGroupToProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroupsPanel()
    End Sub
  
    Protected Sub btnRenameProductGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlProductGroup.SelectedIndex > 0 Then
            lblProductGroupToRename.Text = ddlProductGroup.SelectedItem.Text
            tbRenamedProductGroup.Text = String.Empty
            tbRenamedProductGroup.Focus()
            Call ShowRenameProductGroupPanel()
        Else
            WebMsgBox.Show("Please select a product group to rename")
        End If
    End Sub

    Protected Sub btnAssignPrimaryProductGroupOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlProductGroup.SelectedIndex <= 0 Then
            WebMsgBox.Show("Please select a Product Group")
        Else
            If ddlPrimaryProductGroupOwner.SelectedIndex <= 0 Then
                WebMsgBox.Show("Please select a Primary Product Owner")
            Else
                Call AssignPrimaryProductGroupOwner()
                WebMsgBox.Show(ddlPrimaryProductGroupOwner.SelectedItem.Text & " is now primary owner for product group " & ddlProductGroup.SelectedItem.Text)
            End If
        End If
    End Sub
  
    Protected Sub AssignPrimaryProductGroupOwner()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_UpdatePrimaryProductGroupOwner", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramProductGroupName As SqlParameter = New SqlParameter("@ProductGroupName", SqlDbType.NVarChar, 50)
        paramProductGroupName.Value = ddlProductGroup.SelectedItem.Text
        oCmd.Parameters.Add(paramProductGroupName)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramUserKey As SqlParameter = New SqlParameter("@ProductGroupOwner", SqlDbType.Int, 4)
        paramUserKey.Value = ddlPrimaryProductGroupOwner.SelectedValue
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in AssignPrimaryProductGroupOwner: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnAssignDeputyProductGroupOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlProductGroup.SelectedIndex <= 0 Then
            WebMsgBox.Show("Please select a Product Group")
        Else
            If ddlDeputyProductGroupOwner.SelectedIndex <= 0 Then
                WebMsgBox.Show("Please select a Deputy Product Owner")
            Else
                Call AssignDeputyProductGroupOwner()
                WebMsgBox.Show(ddlDeputyProductGroupOwner.SelectedItem.Text & " is now deputy owner for product group " & ddlProductGroup.SelectedItem.Text)
            End If
        End If
    End Sub
  
    Protected Sub AssignDeputyProductGroupOwner()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_UpdateDeputyProductGroupOwner", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCurrentProductGroupName As SqlParameter = New SqlParameter("@ProductGroupName", SqlDbType.NVarChar, 50)
        paramCurrentProductGroupName.Value = ddlProductGroup.SelectedItem.Text
        oCmd.Parameters.Add(paramCurrentProductGroupName)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramUserKey As SqlParameter = New SqlParameter("@ProductGroupOwner", SqlDbType.Int, 4)
        paramUserKey.Value = ddlDeputyProductGroupOwner.SelectedValue
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in AssignDeputyProductGroupOwner: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub cbUnassignedProductOwnersOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddlPrimaryProductGroupOwner.Items.Clear()
        ddlDeputyProductGroupOwner.Items.Clear()
        Call InitProductGroupControls()
    End Sub
  
    Protected Sub ddlAssignedProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetProductOwners()
    End Sub
  
    Protected Sub SetProductOwners()
        Dim nProductGroup As Integer = ddlAssignedProductGroup.SelectedValue
        Dim nProductOwner1 As Integer
        Dim nProductOwner2 As Integer
        If nProductGroup = 0 Then
            lblAssignedProductOwners.Text = ""
        Else
            ' later refactor this with ddlProductGroup_SelectedIndexChanged
            Dim oDataReader As SqlDataReader = Nothing
            Dim oConn As New SqlConnection(gsConn)
            Dim sSQL As String = "SELECT * FROM ProductGroup WHERE ProductGroupName = '" & ddlAssignedProductGroup.SelectedItem.Text.Replace("'", "''") & "' AND CustomerKey = " & Session("CustomerKey")
            Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                If IsDBNull(oDataReader("ProductOwner1")) Then
                    nProductOwner1 = 0
                Else
                    nProductOwner1 = oDataReader("ProductOwner1")
                End If
                If IsDBNull(oDataReader("ProductOwner2")) Then
                    nProductOwner2 = 0
                Else
                    nProductOwner2 = oDataReader("ProductOwner2")
                End If
            Else
                ' Hmmm, need to work out what to do if read fails
            End If
            lblAssignedProductOwners.Text = "(" & GetProductOwner(nProductOwner1) & " / " & GetProductOwner(nProductOwner2) & ")"
        End If
    End Sub
  
    Protected Sub SetHelpStatus()
        If Request.Cookies("HelpStatus") Is Nothing Then
            Call ShowHelp()
        Else
            Dim sState As String = Request.Cookies("HelpStatus")("CreateNewProduct")
            lnkbtnShowHelp.Text = sState
            If lnkbtnShowHelp.Text = "show help" Then
                Call HideHelp()
            Else
                Call ShowHelp()
            End If
        End If
    End Sub
  
    Protected Sub btn_AddProduct_click(ByVal sender As Object, ByVal e As System.EventArgs)
        pbIsAddingNew = True
        pbIsAddingCategory = False
        pbIsAddingSubCategory = False
        pbIsAddingSubSubCategory = False
        imgbtnDeleteImage.Visible = False
        imgbtnDeletePDF.Visible = False
        btnUploadImage.Visible = False
        btnUploadPDF.Visible = False
        fuBrowseImageFile.Visible = False
        fuBrowsePDFFile.Visible = False
        lblImageUploadUnavailable.Visible = True
        lblPDFUploadUnavailable.Visible = True
        btnAssociatedProducts.Visible = False
        Call SetRequiresAuthAvailability()
        dg_ProductList.CurrentPageIndex = 0
        btnSetUserProfiles.Visible = False
        btn_DeleteProduct.Visible = False
        tbInactivityAlertDays.Text = GetDefaultInactivityAlertDays()
        lnkbtnConfigureCustomLetter.Visible = False
        If IsWURS() Then
            Call GenerateNewFEXCOProductCode()
        End If
        If IsWesternUnion() Then
            Call GenerateNewWesternUnionProductCode()
        End If
        If IsAAT() Then
            Call GenerateNewAATProductCode()
        End If
        Call ShowNewProduct()
        Call SetHelpStatus()
        'End If
    End Sub
  
    Protected Sub lnkbtnShowHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleHelp()
    End Sub
    
    Protected Sub GenerateNewFEXCOProductCode()
        Dim nSeed As Integer = 16
        Dim sProductCode As String, nProductCode As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        If Session("CustomerKey") = CUSTOMER_WURS Then
            sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_WURS & " AND ProductCode LIKE 'WUF/___'"
        Else
            sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_WURS_TEST_ACCOUNT & " AND ProductCode LIKE 'WUF/___'"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read()
                    sProductCode = oDataReader("ProductCode").ToString.Substring(4, 3)
                    If IsNumeric(sProductCode) Then
                        nProductCode = CInt(sProductCode)
                        If nProductCode > nSeed Then
                            nSeed = nProductCode
                        End If
                    End If
                End While
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
        nSeed += 1
        txtProductCode.Text = "WUF/" & Format(nSeed, "000")
    End Sub
  
    Protected Sub GenerateNewWesternUnionProductCode()
        Dim nSeed As Integer = 0
        Dim sProductCode As String, nProductCode As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_WESTERN_UNION & " AND ProductCode LIKE 'WU/____'"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read()
                    sProductCode = oDataReader("ProductCode")
                    If sProductCode.Length >= 7 Then
                        sProductCode = sProductCode.Substring(3, 4)
                        If IsNumeric(sProductCode) Then
                            nProductCode = CInt(sProductCode)
                            If nProductCode > nSeed Then
                                nSeed = nProductCode
                            End If
                        End If
                    End If
                End While
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
        nSeed += 1
        txtProductCode.Text = "WU/" & Format(nSeed, "0000")
    End Sub

    Protected Sub GenerateNewAATProductCode()
        Dim nSeed As Integer = 0
        Dim sProductCode As String, nProductCode As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & CUSTOMER_AAT & " AND ProductCode LIKE 'AAT/______'"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read()
                    sProductCode = oDataReader("ProductCode")
                    If sProductCode.Length >= 10 Then
                        sProductCode = sProductCode.Substring(4, 6)
                        If IsNumeric(sProductCode) Then
                            nProductCode = CInt(sProductCode)
                            If nProductCode > nSeed Then
                                nSeed = nProductCode
                            End If
                        End If
                    End If
                End While
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
        nSeed += 1
        txtProductCode.Text = "AAT/" & Format(nSeed, "000000")
    End Sub

    Protected Sub StoreHelpStatus()
        Dim c As HttpCookie = New HttpCookie("HelpStatus")
        c.Values.Add("CreateNewProduct", lnkbtnShowHelp.Text)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
        Response.Flush()
    End Sub

    Protected Sub ToggleHelp()
        If lnkbtnShowHelp.Text = "show help" Then
            lnkbtnShowHelp.Text = "hide help"
            Call ShowHelp()
        Else
            lnkbtnShowHelp.Text = "show help"
            Call HideHelp()
        End If
        Call StoreHelpStatus()
        Call CheckVisibility()
    End Sub
  
    Protected Sub SetVisible(ByVal a As HtmlContainerControl)
        If a.InnerText <> "NULL" Then
            a.Visible = True
        End If
    End Sub
   
    Protected Sub ShowHelp()
        If Not pbIsAddingNew Then
            aHelpDeleteProduct.Visible = True
        End If
        Call SetVisible(aHelpProductCode)
        Call SetVisible(aHelpProductDate)
        Call SetVisible(aHelpMinStockLevel)
        Call SetVisible(aHelpDescription)
        Call SetVisible(aHelpItemsPerBox)
        Call SetVisible(aHelpCategory)
        Call SetVisible(aHelpSubCategory)
        Call SetVisible(aHelpArchived)
        Call SetVisible(aHelpSubSubCategory)
        Call SetVisible(aHelpInactivityAlert)
        Call SetVisible(aHelpCalendarManaged)

        Call SetVisible(aHelpLanguage)
        Call SetVisible(aHelpProspectusNumbers)
        Call SetVisible(aHelpViewOnWebForm)
        Call SetVisible(aHelpAdRotatorText)
        Call SetVisible(aHelpAdRotator)
        Call SetVisible(aHelpDepartment)
        Call SetVisible(aHelpUnitWeight)
        Call SetVisible(aHelpMisc1)
        Call SetVisible(aHelpMisc2)
        Call SetVisible(aHelpUnitValue)
        Call SetVisible(aHelpAssignedProductGroup)
        Call SetVisible(aHelpExpiryDate)
        Call SetVisible(aHelpReplenishmentDate)
        Call SetVisible(aHelpNotes)
        Call SetVisible(aHelpUploadImage)
        Call SetVisible(aHelpUploadPDF)
        Call SetVisible(aHelpCustomLetter)
        Call SetVisible(aHelpSellingPrice)
        Call SetVisible(aHelpProductCredits)
    End Sub
  
    Protected Sub HideHelp()
        aHelpDeleteProduct.Visible = False
        aHelpProductCode.Visible = False
        aHelpProductDate.Visible = False
        aHelpMinStockLevel.Visible = False
        aHelpDescription.Visible = False
        aHelpItemsPerBox.Visible = False
        aHelpCategory.Visible = False
        aHelpSubCategory.Visible = False
        aHelpArchived.Visible = False
        aHelpSubSubCategory.Visible = False
        aHelpInactivityAlert.Visible = False
        aHelpCalendarManaged.Visible = False
        aHelpLanguage.Visible = False
        aHelpProspectusNumbers.Visible = False
        aHelpViewOnWebForm.Visible = False
        aHelpAdRotatorText.Visible = False
        aHelpAdRotator.Visible = False
        aHelpDepartment.Visible = False
        aHelpUnitWeight.Visible = False
        aHelpMisc1.Visible = False
        aHelpMisc2.Visible = False
        aHelpUnitValue.Visible = False
        aHelpAssignedProductGroup.Visible = False
        aHelpExpiryDate.Visible = False
        aHelpReplenishmentDate.Visible = False
        aHelpNotes.Visible = False
        aHelpUploadImage.Visible = False
        aHelpUploadPDF.Visible = False
        aHelpCustomLetter.Visible = False
        aHelpSellingPrice.Visible = False
        aHelpProductCredits.Visible = False
    End Sub
  
    Protected Sub lnkbtnNewProductCode_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProductCode.Text = FormatNextFreeProductCode()
        lnkbtnNewProductCode.Visible = False
    End Sub
  
    Protected Function FormatNextFreeProductCode() As String
        Dim sCode As String
        Do
            sCode = CStr(GetNextFreeProductCode())
            While sCode.Length < 6
                sCode = "0" & sCode
            End While
            sCode = "W" & sCode
        Loop Until IsUniqueProductCode(sCode, Session("CustomerKey"))
        FormatNextFreeProductCode = sCode
    End Function

    Protected Function GetNextFreeProductCode() As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        Dim nCode As Integer
        sSQL = "SELECT * FROM LANDGControl"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        GetNextFreeProductCode = 0
        If oDataReader.HasRows Then
            oDataReader.Read()
            nCode = oDataReader("ProductCodeSeed")
        End If
        oDataReader.Close()
        GetNextFreeProductCode = nCode
        nCode += 1
        sSQL = "UPDATE LANDGControl SET ProductCodeSeed = " & nCode
        Dim oCmd2 As SqlCommand = New SqlCommand(sSQL, oConn)
        oCmd2.ExecuteNonQuery()
        oConn.Close()
    End Function

    Protected Function IsUniqueProductCode(ByVal sProductCode As String, ByVal nCustomerKey As Integer) As Boolean
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT * FROM LogisticProduct WHERE ProductCode = '" & sProductCode & "' AND CustomerKey = " & nCustomerKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        If oDataReader.HasRows Then
            IsUniqueProductCode = False
        Else
            IsUniqueProductCode = True
        End If
        oConn.Close()
    End Function
  
    Protected Sub ddlStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = 4 Then
            chkArchivedFlag.Checked = True
        End If
    End Sub
  
    Protected Sub GetPendingOrderAuthorisations()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_AuthOrderGetPendingRequests", oConn)
      
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@AuthoriserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@AuthoriserKey").Value = Session("UserKey")

        oAdapter.Fill(oDatatable)
        gvAuthoriseOrder.DataSource = oDatatable
        gvAuthoriseOrder.DataBind()
    End Sub
  
    Protected Function GetOrderAuthorisationByKey(ByVal nHoldingQueueKey As Integer) As DataRow
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_AuthOrderGetByKey", oConn)
        GetOrderAuthorisationByKey = Nothing
        Try
          
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@HoldingQueueKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@HoldingQueueKey").Value = nHoldingQueueKey
            oAdapter.Fill(oDatatable)
            GetOrderAuthorisationByKey = oDatatable.Rows(0)
        Catch ex As Exception
            WebMsgBox.Show("Internal error in GetOrderAuthorisationByKey - " & ex.ToString)
        Finally
            oConn.Close()
        End Try
    End Function
  
    Protected Sub btnShowAuthOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnShowAuthOrder As Button = sender
        Dim nHoldingQueueKey As Integer = btnShowAuthOrder.CommandArgument
        Call GetAuthorisationOrder(nHoldingQueueKey)
        Call ShowAuthOrderDetailsPanel()
    End Sub
  
    Protected Function GetAuthOrderDetails(ByVal nHoldingQueueKey As Integer) As DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_AuthOrderGetDetails", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@HoldingQueueKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@HoldingQueueKey").Value = nHoldingQueueKey
            oAdapter.Fill(oDatatable)
        Catch ex As Exception
            WebMsgBox.Show("Internal error in GetAuthorisationOrder - " & ex.ToString)
        Finally
            GetAuthOrderDetails = oDatatable
            oConn.Close()
        End Try
    End Function
  
    Protected Sub GetAuthorisationOrder(ByVal nHoldingQueueKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As DataTable = GetAuthOrderDetails(nHoldingQueueKey)
        Dim drOrderDetails As DataRow = GetOrderAuthorisationByKey(nHoldingQueueKey)
        lblAuthOrderOrderedBy.Text = drOrderDetails.Item("FirstName") & " " & drOrderDetails.Item("LastName")
        lblAuthOrderPlacedOn.Text = drOrderDetails.Item("OrderCreatedDateTime")
        lblAuthMsgToAuthoriser.Text = drOrderDetails.Item("MsgToAuthoriser")
        lblAuthOrderConsignee.Text = drOrderDetails.Item("CneeName")
        lblAuthOrderAttnOf.Text = drOrderDetails.Item("CneeCtcName")
        lblAuthOrderAddr1.Text = drOrderDetails.Item("CneeAddr1")
        lblAuthOrderAddr2.Text = drOrderDetails.Item("CneeAddr2")
        lblAuthOrderAddr3.Text = drOrderDetails.Item("CneeAddr3")
        lblAuthOrderTown.Text = drOrderDetails.Item("CneeTown")
        lblAuthOrderState.Text = drOrderDetails.Item("CneeState")
        lblAuthOrderPostcode.Text = drOrderDetails.Item("CneePostCode")
        lblAuthOrderCountry.Text = drOrderDetails.Item("CountryName")
        hidHoldingQueueKey.Value = nHoldingQueueKey
        gvAuthOrderDetails.DataSource = oDatatable
        gvAuthOrderDetails.DataBind()
    End Sub
  
    Protected Function CheckValidOrder() As String
        Dim lblAuthProductCode As Label
        Dim hidQtyAvailable As HiddenField
        Dim hidArchiveFlag As HiddenField
        Dim hidDeletedFlag As HiddenField
        Dim tbAuthOrderQty As TextBox
        Dim nQtyRequired As Integer
        Dim sbResult As New StringBuilder
        Dim bNonZeroQtyFound As Boolean = False
        For Each gvr As GridViewRow In gvAuthOrderDetails.Rows
            lblAuthProductCode = gvr.Cells(0).FindControl("lblAuthProductCode")
            hidQtyAvailable = gvr.Cells(0).FindControl("hidQtyAvailable")
            hidArchiveFlag = gvr.Cells(0).FindControl("hidArchiveFlag")
            hidDeletedFlag = gvr.Cells(0).FindControl("hidDeletedFlag")
            tbAuthOrderQty = gvr.Cells(3).FindControl("tbAuthOrderQty")
            nQtyRequired = IsBlankOrPositiveInteger(tbAuthOrderQty.Text)
            If nQtyRequired > 0 Then
                bNonZeroQtyFound = True
                If hidArchiveFlag.Value <> "N" Then
                    sbResult.Append("Product " & lblAuthProductCode.Text & " is archived. Archived products cannot be ordered.")
                    sbResult.Append("\n")
                ElseIf hidDeletedFlag.Value <> "N" Then
                    sbResult.Append("Product " & lblAuthProductCode.Text & " is deleted. Deleted products cannot be ordered.")
                    sbResult.Append("\n")
                ElseIf hidQtyAvailable.Value = "0" Then
                    sbResult.Append("Product " & lblAuthProductCode.Text & " has no stock available.")
                    sbResult.Append("\n")
                ElseIf nQtyRequired > CInt(hidQtyAvailable.Value) Then
                    sbResult.Append("Product " & lblAuthProductCode.Text & " has insufficient stock quantity (" & hidQtyAvailable.Value & ") to fulfil this order.")
                    sbResult.Append("\n")
                End If
            Else
                If nQtyRequired = -1 Then
                    sbResult.Append("Product " & lblAuthProductCode.Text & " has unrecognised quantity value")
                    sbResult.Append("\n")
                End If
            End If
        Next
        If sbResult.Length = 0 AndAlso Not bNonZeroQtyFound Then
            sbResult.Append("There appear to be no items in this order. You must either add items or decline authorisation.")
        End If
        CheckValidOrder = sbResult.ToString
    End Function
  
    Protected Sub btnOrderAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call OrderAuthorise()
    End Sub
  
    Protected Sub OrderAuthorise()
        Dim sValidationResult As String = CheckValidOrder()
        If sValidationResult.Length > 0 Then
            WebMsgBox.Show(sValidationResult)
        Else
            Dim lConsignmentKey As Long = SubmitOrder()
            If lConsignmentKey > 0 Then
                Dim sMessage As String
                Call UpdateHoldingQueueEntry("COMPLETE", lConsignmentKey)
                Call EmailOrderer(bSuccess:=True, lConsignmentKey:=lConsignmentKey)
                Call GetPendingOrderAuthorisations()
                sMessage = "Authorisation complete. The consignment number for that order is " & lConsignmentKey.ToString & "."
                If gvAuthoriseOrder.Rows.Count = 0 Then
                    sMessage += "\n\nYou have no further authorisation requests awaiting approval."
                    Call ShowMainPanel()
                Else
                    Call ShowAuthoriseOrderPanel()
                End If
                WebMsgBox.Show(sMessage)
            ElseIf lConsignmentKey = 0 Then
                WebMsgBox.Show("Internal error during authorisation")
            ElseIf lConsignmentKey = -1 Then
                WebMsgBox.Show("This order has already been processed")
            End If
        End If
    End Sub

    Protected Sub EmailOrderer(ByVal bSuccess As Boolean, ByVal lConsignmentKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderEmailOrderer", oConn)
        Dim spParam As SqlParameter
        oCmd.CommandType = CommandType.StoredProcedure

        spParam = New SqlParameter("@StatusFlag", SqlDbType.Bit)
        If bSuccess Then
            spParam.Value = 1
        Else
            spParam.Value = 0
        End If
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@ConsignmentKey", SqlDbType.NVarChar, 50)
        spParam.Value = lConsignmentKey
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@HoldingQueueKey", SqlDbType.Int)
        spParam.Value = hidHoldingQueueKey.Value
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@Message", SqlDbType.NVarChar, 1000)
        spParam.Value = tbAuthOrderMessage.Text
        oCmd.Parameters.Add(spParam)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.ToString)
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub UpdateHoldingQueueEntry(ByVal sStatus As String, ByVal lConsignmentKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderUpdateHoldingQueue", oConn)
        Dim spParam As SqlParameter
        oCmd.CommandType = CommandType.StoredProcedure

        spParam = New SqlParameter("@HoldingQueueKey", SqlDbType.Int)
        spParam.Value = hidHoldingQueueKey.Value
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@OrderStatus", SqlDbType.NVarChar, 50)
        spParam.Value = sStatus
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@ConsignmentKey", SqlDbType.Int)
        spParam.Value = lConsignmentKey
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@MsgToOrderer", SqlDbType.NVarChar, 1000)
        spParam.Value = tbAuthOrderMessage.Text.Replace(Environment.NewLine, " ")
        oCmd.Parameters.Add(spParam)
      
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.ToString)
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub btnOrderDecline_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sMessage As String
        Dim drOrderDetails As DataRow = GetOrderAuthorisationByKey(hidHoldingQueueKey.Value)
        If Not drOrderDetails("OrderStatus") = "QUEUED" Then
            WebMsgBox.Show("This order has already been processed")
            Call ShowMainPanel()
            Exit Sub
        End If

        Call UpdateHoldingQueueEntry("DECLINED", 0)
        Call EmailOrderer(bSuccess:=False, lConsignmentKey:=0)
        sMessage = "Authorisation declined."
        Call GetPendingOrderAuthorisations()
        If gvAuthoriseOrder.Rows.Count = 0 Then
            sMessage += "\n\nYou have no further authorisation requests awaiting approval."
            Call ShowMainPanel()
        Else
            Call ShowAuthoriseOrderPanel()
        End If
        WebMsgBox.Show(sMessage)
    End Sub

    Protected Function IsBlankOrPositiveInteger(ByVal sString As String) As Integer
        IsBlankOrPositiveInteger = -1
        sString = sString.Trim
        If sString.Length = 0 Then
            Return 0
        End If
        If Not IsNumeric(sString) Then
            Exit Function
        End If
        For Each c As Char In sString
            If Not Char.IsDigit(c) Then
                Return -1
            End If
        Next
        Return CInt(sString)
    End Function
  
    Protected Function SubmitOrder() As Long
        SubmitOrder = 0
        Dim drOrderDetails As DataRow = GetOrderAuthorisationByKey(hidHoldingQueueKey.Value)
        If Not drOrderDetails("OrderStatus") = "QUEUED" Then
            SubmitOrder = -1
            Exit Function
        End If

        Dim sSpecialInstr As String
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
      
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = drOrderDetails("UserProfileKey")
        oCmdAddBooking.Parameters.Add(param1)
      
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = drOrderDetails("CustomerKey")
        oCmdAddBooking.Parameters.Add(param2)
      
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = drOrderDetails("BookingReference1")
        oCmdAddBooking.Parameters.Add(param3)
      
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = drOrderDetails("BookingReference2")
        oCmdAddBooking.Parameters.Add(param4)
      
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = drOrderDetails("BookingReference3")
        oCmdAddBooking.Parameters.Add(param5)
      
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = drOrderDetails("BookingReference4")
        oCmdAddBooking.Parameters.Add(param6)
          
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = drOrderDetails("ExternalReference")
        oCmdAddBooking.Parameters.Add(param6a)
      
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        sSpecialInstr = drOrderDetails("SpecialInstructions")
        sSpecialInstr = Replace(sSpecialInstr, vbCrLf, " ")
        param7.Value = sSpecialInstr
        oCmdAddBooking.Parameters.Add(param7)
      
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = drOrderDetails("PackingNoteInfo")
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
        param13.Value = drOrderDetails("CnorName")
        oCmdAddBooking.Parameters.Add(param13)
      
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = drOrderDetails("CnorAddr1")
        oCmdAddBooking.Parameters.Add(param14)
      
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = drOrderDetails("CnorAddr2")
        oCmdAddBooking.Parameters.Add(param15)
      
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = drOrderDetails("CnorAddr3")
        oCmdAddBooking.Parameters.Add(param16)
      
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = drOrderDetails("CnorTown")
        oCmdAddBooking.Parameters.Add(param17)
      
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = drOrderDetails("CnorState")
        oCmdAddBooking.Parameters.Add(param18)
      
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = drOrderDetails("CnorPostCode")
        oCmdAddBooking.Parameters.Add(param19)
      
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = CLng(drOrderDetails("CnorCountryKey"))
        oCmdAddBooking.Parameters.Add(param20)
      
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = drOrderDetails("CnorCtcName")
        oCmdAddBooking.Parameters.Add(param21)
      
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = drOrderDetails("CnorTel")
        oCmdAddBooking.Parameters.Add(param22)
      
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = drOrderDetails("CnorEmail")
        oCmdAddBooking.Parameters.Add(param23)
      
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
      
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = drOrderDetails("CneeName")
        oCmdAddBooking.Parameters.Add(param25)
      
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = drOrderDetails("CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)
      
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = drOrderDetails("CneeAddr2")
        oCmdAddBooking.Parameters.Add(param27)
      
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = drOrderDetails("CneeAddr3")
        oCmdAddBooking.Parameters.Add(param28)
      
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = drOrderDetails("CneeTown")
        oCmdAddBooking.Parameters.Add(param29)
      
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = drOrderDetails("CneeState")
        oCmdAddBooking.Parameters.Add(param30)
      
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = drOrderDetails("CneePostCode")
        oCmdAddBooking.Parameters.Add(param31)
      
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = drOrderDetails("CneeCountryKey")
        oCmdAddBooking.Parameters.Add(param32)
      
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = drOrderDetails("CneeCtcName")
        oCmdAddBooking.Parameters.Add(param33)
      
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = drOrderDetails("CneeTel")
        oCmdAddBooking.Parameters.Add(param34)
      
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = drOrderDetails("CneeEmail")
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
                If gvAuthOrderDetails.Rows.Count > 0 Then
                    For Each gvr As GridViewRow In gvAuthOrderDetails.Rows
                        Dim hidLogisticProductKey As HiddenField = gvr.Cells(0).FindControl("hidLogisticProductKey")
                        Dim tbAuthOrderQty As TextBox = gvr.Cells(3).FindControl("tbAuthOrderQty")
                        Dim lProductKey As Long = CLng(hidLogisticProductKey.Value)
                        Dim lPickQuantity As Long = CLng(tbAuthOrderQty.Text)
                        If lPickQuantity > 0 Then
                            Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                            oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                          
                            Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                            param51.Value = CLng(drOrderDetails("UserProfileKey"))
                            oCmdAddStockItem.Parameters.Add(param51)
                          
                            Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                            param52.Value = CLng(drOrderDetails("CustomerKey"))
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
                        End If
                        Dim oCmdUpdateAuthorisedQuantity As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderUpdateItemHoldingQueue", oConn)
                        oCmdUpdateAuthorisedQuantity.CommandType = CommandType.StoredProcedure
                      
                        Dim param60 As SqlParameter = New SqlParameter("@OrderHoldingQueueKey", SqlDbType.Int, 4)
                        param60.Value = CInt(hidHoldingQueueKey.Value)
                        oCmdUpdateAuthorisedQuantity.Parameters.Add(param60)

                        Dim param61 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param61.Value = lProductKey
                        oCmdUpdateAuthorisedQuantity.Parameters.Add(param61)

                        Dim param62 As SqlParameter = New SqlParameter("@ItemsOutAuthorised", SqlDbType.Int, 4)
                        param62.Value = lPickQuantity
                        oCmdUpdateAuthorisedQuantity.Parameters.Add(param62)

                        oCmdUpdateAuthorisedQuantity.Connection = oConn
                        oCmdUpdateAuthorisedQuantity.Transaction = oTrans
                        oCmdUpdateAuthorisedQuantity.ExecuteNonQuery()
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
                    lblError.Text = "No stock items found for booking"
                End If
            Else
                BookingFailed = True
                lblError.Text = "Error adding Web Booking [BookingKey=0]."
            End If
            If Not BookingFailed Then
                oTrans.Commit()
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
        If Not BookingFailed Then
            SubmitOrder = lConsignmentKey
        End If
    End Function
  
    Protected Function gvAuthOrderDetailsItemForeColor(ByVal DataItem As Object) As System.Drawing.Color
        gvAuthOrderDetailsItemForeColor = Black
        If Not IsDBNull(DataBinder.Eval(DataItem, "Authorised")) AndAlso DataBinder.Eval(DataItem, "Authorised") = "N" Then
            gvAuthOrderDetailsItemForeColor = Red
        End If
    End Function

    Protected Sub dgPreAuthorise_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dgi As DataGridItem = e.Item
        If dgi.ItemType = ListItemType.Item Or dgi.ItemType = ListItemType.AlternatingItem Then
            Dim sKey As String = dgi.Cells(0).Text

            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Product_GetAuthorised", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam1 As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int, 4)
            oParam1.Value = plProductKey
            Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@UserProfileKey", SqlDbType.Int, 4)
            oParam2.Value = sKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                If oDataReader.HasRows Then
                    oDataReader.Read()
                    Dim sAuthorisedQuantity As String = oDataReader("AuthorisedQuantity")
                    Dim sQuantityRemaining As String = oDataReader("QuantityRemaining")
                    Dim l As Label = dgi.Cells(4).Controls(1)
                    l.ForeColor = Red
                    l.Text = "(" & sQuantityRemaining & " unused of " & sAuthorisedQuantity & " auth'd)"
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub

    Protected Sub btnSaveNewAuthoriser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveNewAuthoriser()
        cbRequiresAuth.Checked = True
        lnkbtnPreAuthorise.Visible = True
        Call ShowProductDetail()
    End Sub
  
    Protected Sub SaveNewAuthoriser()
        Dim oConn As New SqlConnection(gsConn)
        Dim sProc As String
        If Not (cbChangeAuthoriserOnAllSelective.Checked Or cbChangeAuthoriserOnAll.Checked) Then
            sProc = "spASPNET_Product_UpdateAuthoriser2"
        ElseIf cbChangeAuthoriserOnAll.Checked Then
            sProc = "spASPNET_Product_UpdateAuthoriserAllProducts2"
        Else
            sProc = "spASPNET_Product_UpdateAuthoriserAllMatchingProducts2"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sProc, oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)

        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int)
        paramUserProfileKey.Value = ddlModifyAuthoriser.SelectedValue
        oCmd.Parameters.Add(paramUserProfileKey)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.ToString)
        Finally
            oConn.Close()
        End Try
    End Sub
  
    Protected Sub ddlModifyAuthoriser_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlModifyAuthoriser.SelectedIndex > 0 Then
            btnSaveNewAuthoriser.Enabled = True
        Else
            btnSaveNewAuthoriser.Enabled = False
        End If
    End Sub
  
    Protected Sub imgbtnDeleteImage_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Call DeleteImage()
    End Sub
  
    Protected Sub DeleteImage()
        Dim SavePath As String = psProdThumbFolder & plProductKey.ToString & ".jpg"
        If System.IO.File.Exists(SavePath) Then
            System.IO.File.Delete(SavePath)
        End If
        Dim sSavePath As String = psProdImageFolder & plProductKey.ToString & ".jpg"
        If System.IO.File.Exists(sSavePath) Then
            System.IO.File.Delete(sSavePath)
        End If
        Call ResetImageAttributes()
        hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & "blank_thumb.jpg"
        hlnk_DetailThumb.NavigateUrl = psVirtualJPGFolder & "blank_image.jpg"
        imgbtnDeleteImage.Visible = False
    End Sub

    Protected Sub imgbtnDeletePDF_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Call DeletePDF()
    End Sub
  
    Protected Sub DeletePDF()
        hlnk_PDF.ImageUrl = psVirtualPDFFolder & "blank_pdf_thumb.jpg"
        hlnk_PDF.NavigateUrl = psVirtualPDFFolder & "blank_pdf.jpg"
        Call SetPDFAttribute()
        imgbtnDeletePDF.Visible = False
    End Sub
  
    Protected Sub btnAssociatedProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAssociatedProductsPanel()
    End Sub

    Protected Sub InitAssociatedProductsPanel()
        Call PopulateAssociatedProductsGrid()
        Call PopulateUnassociatedProductsGrid()
        If txtProductDate.Text.Trim.Length > 0 Then
            lblAssociatedProductsProductCode.Text = txtProductCode.Text & "-" & txtProductDate.Text
        Else
            lblAssociatedProductsProductCode.Text = txtProductCode.Text
        End If
    End Sub
   
    Protected Sub PopulateAssociatedProductsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetAssociatedProducts", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = plProductKey

            oAdapter.Fill(oDataTable)
            gvAssociatedProducts.DataSource = oDataTable
            gvAssociatedProducts.DataBind()
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Sub
   
    Protected Sub PopulateUnassociatedProductsGrid()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetUnassociatedProducts", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = plProductKey

            oAdapter.Fill(oDataTable)
            gvUnassociatedProducts.DataSource = oDataTable
            gvUnassociatedProducts.DataBind()
        Catch ex As SqlException
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnAddAssociatedProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nAssociatedProductKey As Integer = b.CommandArgument
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AddAssociatedProduct", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        Dim paramAssociatedProductKey As SqlParameter = New SqlParameter("@AssociatedProductKey", SqlDbType.Int)
        paramAssociatedProductKey.Value = nAssociatedProductKey
        oCmd.Parameters.Add(paramAssociatedProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call InitAssociatedProductsPanel()
    End Sub

    Protected Sub btnRemoveAssociatedProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLogisticAssociatedProductKey As Integer = b.CommandArgument
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product__Product_DeleteAssociatedProduct", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramLogisticAssociatedProductKey As SqlParameter = New SqlParameter("@LogisticAssociatedProductKey", SqlDbType.Int)
        paramLogisticAssociatedProductKey.Value = nLogisticAssociatedProductKey
        oCmd.Parameters.Add(paramLogisticAssociatedProductKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        Call InitAssociatedProductsPanel()
    End Sub
   
    Protected Sub cbCalendarManaged_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            If cbRequiresAuth.Checked Then
                WebMsgBox.Show("Cannot set the Calendar Managed attribute on an authorisable product. Only one of these facilities can be used. Remove the Requires Authorisation attribute if you want to use Calendar Management.")
                cb.Checked = False
            End If
        End If
        If cb.Checked Then
            lblLegendLanguage.Text = "Type:"
        Else
            lblLegendLanguage.Text = "Language:"
        End If
    End Sub
   
    Protected Sub btnShowProductsInGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductsInGroup()
    End Sub
   
    Protected Sub ShowProductsInGroup()
        Dim sSQL As String = "SELECT ProductCode 'Product Code', ProductDate 'Value / Date', ProductDescription 'Description' FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " AND StockOwnedByKey = " & ddlProductGroup.SelectedValue

        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("ShowProductsInGroup: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        gvProductsInGroup.DataSource = oDataTable
        gvProductsInGroup.DataBind()
        gvProductsInGroup.Visible = True
    End Sub
   
    Protected Sub lnkbtnConfigureProductInactivityAlerts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Dim nDefaultInactivityAlertDays As Integer = GetDefaultInactivityAlertDays()
        tbProductInactivityAlertPeriodExistingProducts.Text = nDefaultInactivityAlertDays
        tbProductInactivityAlertPeriodNewProducts.Text = nDefaultInactivityAlertDays
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            btnSetProductInactivityAlertNewProducts.Enabled = False
            tbProductInactivityAlertPeriodNewProducts.Enabled = False
            lblAvailableToSuperUsersOnly.Visible = True
        End If
        pnlConfigureProductInactivityAlert.Visible = True
    End Sub

    Protected Sub btnViewProductsUsingInactivityAlert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call ViewProductsUsingInactivityAlert()
        pnlProductInactivityAlertStatus.Visible = True
    End Sub
    
    Protected Sub ViewProductsUsingInactivityAlert()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectFields As String = "SELECT ProductCode 'Product Code', ProductDate 'Value/Date', ProductDescription 'Description', InactivityAlertDays 'Inactivity Alert Days' FROM LogisticProduct lp "
        Dim sCondition As String = " AND ISNULL(InactivityAlertDays,0) > 0 ORDER BY LogisticProductKey"
        Dim sSQL As String
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sSQL = sSelectFields & "INNER JOIN ProductGroup pg ON lp.StockOwnedByKey = pg.ProductGroupKey WHERE pg.ProductOwner1 = " & Session("UserKey") & " OR pg.ProductOwner2 = " & Session("UserKey")
        Else
            sSQL = sSelectFields & "WHERE CustomerKey = " & Session("CustomerKey") & sCondition
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            gvProductInactivityAlert.DataSource = oDataReader
            gvProductInactivityAlert.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("Error in ViewProductsUsingInactivityAlert: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetDefaultInactivityAlertDays() As Integer
        GetDefaultInactivityAlertDays = 0
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "SELECT DefaultInactivityAlertDays FROM Customer WHERE CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader(0)) Then
                GetDefaultInactivityAlertDays = oDataReader(0)
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetDefaultInactivityAlertDays: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub btnBackFromConfigureProductInactivityAlert_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Sub btnBackFromProductInactivityAlertStatus_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Sub btnSetProductInactivityAlertAllExistingProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetProductInactivityAlertAllExistingProducts()
    End Sub

    Protected Sub SetProductInactivityAlertAllExistingProducts()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectField As String = "SELECT LogisticProductKey FROM LogisticProduct "
        Dim sSQL As String
        Dim nResult As Integer
        If Session("UserType").ToString.ToLower.Contains("owner") Then
            sSQL = sSelectField & "INNER JOIN ProductGroup pg ON lp.StockOwnedByKey = pg.ProductGroupKey WHERE pg.ProductOwner1 = " & Session("UserKey") & " OR pg.ProductOwner2 = " & Session("UserKey")
        Else
            sSQL = sSelectField & "WHERE CustomerKey = " & Session("CustomerKey")
        End If
        sSQL = "UPDATE LogisticProduct SET InactivityAlertDays = " & tbProductInactivityAlertPeriodExistingProducts.Text & " WHERE LogisticProductKey IN (" & sSQL & ")"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            nResult = oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetProductInactivityAlertAllExistingProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Dim sProducts As String = " product"
        If nResult <> 1 Then
            sProducts += "s"
        End If
        Dim sDays As String = " day"
        If CInt(tbProductInactivityAlertPeriodExistingProducts.Text) > 1 Then
            sDays += "s"
            sDays += "."
        End If
        If CInt(tbProductInactivityAlertPeriodExistingProducts.Text) > 0 Then
            WebMsgBox.Show(nResult.ToString & sProducts & " assigned an Inactivity Alert value of " & tbProductInactivityAlertPeriodExistingProducts.Text & " day(s)")
        Else
            WebMsgBox.Show(nResult.ToString & sProducts & " will no longer receive Inactivity Alerts")
        End If
        If Not pbIsAddingNew Then
            tbInactivityAlertDays.Text = tbProductInactivityAlertPeriodExistingProducts.Text
        End If
    End Sub
    
    Protected Sub btnSetProductInactivityAlertNewProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetProductInactivityAlertNewProducts()
    End Sub
    
    Protected Sub SetProductInactivityAlertNewProducts()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSelectField As String = "SELECT LogisticProductKey FROM LogisticProduct "
        Dim sSQL As String
        sSQL = "UPDATE Customer SET DefaultInactivityAlertDays = " & tbProductInactivityAlertPeriodNewProducts.Text & " WHERE CustomerKey = " & Session("CustomerKey")
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("Error in SetProductInactivityAlertAllExistingProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Dim sDays As String = " day"
        If CInt(tbProductInactivityAlertPeriodNewProducts.Text) > 1 Then
            sDays += "s"
        End If
        sDays += "."
        If CInt(tbProductInactivityAlertPeriodNewProducts.Text) > 0 Then
            WebMsgBox.Show("Newly created products will be assigned an Inactivity Alert value of " & tbProductInactivityAlertPeriodNewProducts.Text & sDays)
        Else
            WebMsgBox.Show("Newly created products will not receive Inactivity Alerts.")
        End If
        If pbIsAddingNew Then
            tbInactivityAlertDays.Text = tbProductInactivityAlertPeriodNewProducts.Text
        End If
    End Sub
    
    Protected Sub lnkbtnConfigureCustomLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitCustomLetterPanel()
        Call ShowConfigureCustomLetterPanel()
    End Sub
    
    Protected Sub InitCustomLetterPanel()
        
    End Sub
    
    Protected Sub cbCustomLetter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If Not pbIsAddingNew Then
            lnkbtnConfigureCustomLetter.Visible = cb.Checked
        End If
    End Sub

    'Protected Sub ShowNewKODDFISProduct()
    '    Call HideAllPanels()
    '    pnlNewKODDFISProduct.Visible = True
    '    Call GetKODDFISCategories()
    '    ddlKODDFISSubCategory.Visible = False
    '    lblLegendKODDFISSubCategory.Visible = False
    'End Sub

    'Protected Sub GetKODDFISCategories()
    '    Dim oDataReader As SqlDataReader
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim sSQL As String = "SELECT DISTINCT Category FROM ClientData_KODDFIS_Categories ORDER BY Category"
    '    Dim oCmd As New SqlCommand(sSQL, oConn)
    '    ddlKODDFISCategory.Items.Clear()
    '    ddlKODDFISCategory.Items.Add(New ListItem("- please select -", 0))
    '    Try
    '        oConn.Open()
    '        oDataReader = oCmd.ExecuteReader()
    '        While oDataReader.Read()
    '            ddlKODDFISCategory.Items.Add(New ListItem(oDataReader("Category")))
    '        End While
    '    Catch ex As SqlException
    '        WebMsgBox.Show("Error in GetKODDFISCategories: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Sub
    
    'Protected Sub ddlKODDFISCategory_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim ddl As DropDownList = sender
    '    If ddl.SelectedIndex > 0 Then
    '        If ddlKODDFISCategory.Items(0).Text = "- please select -" Then
    '            ddlKODDFISCategory.Items.RemoveAt(0)
    '        End If
    '        ddlKODDFISSubCategory.Visible = True
    '        lblLegendKODDFISSubCategory.Visible = True
    '        Call GetKODDFISSubCategories()
    '    End If
    'End Sub

    'Protected Sub GetKODDFISSubCategories()
    '    Dim oDataReader As SqlDataReader
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim sSQL As String = "SELECT SubCategory FROM ClientData_KODDFIS_Categories WHERE Category = '" & ddlKODDFISCategory.SelectedItem.Text & "' ORDER BY SubCategory"
    '    Dim oCmd As New SqlCommand(sSQL, oConn)
    '    ddlKODDFISSubCategory.Items.Clear()
    '    ddlKODDFISSubCategory.Items.Add(New ListItem("- please select -", 0))
    '    Try
    '        oConn.Open()
    '        oDataReader = oCmd.ExecuteReader()
    '        While oDataReader.Read()
    '            ddlKODDFISSubCategory.Items.Add(New ListItem(oDataReader("SubCategory")))
    '        End While
    '    Catch ex As SqlException
    '        WebMsgBox.Show("Error in GetKODDFISSubCategories: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Sub
    
    'Protected Sub btnCreateNewKODDFISProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call CreateNewKODDFISProduct()
    'End Sub
    
    'Protected Sub CreateNewKODDFISProduct()
    '    Page.Validate("KODDFIS")
    '    If Page.IsValid Then
    '        Dim sProductCode As String = GetKODDFISDivisionCode()
    '        sProductCode += GetNextKODDFISCode(sProductCode)
    '        txtProductCode.Text = sProductCode
            
    '        txtCategory.Visible = True
    '        txtCategory.Text = ddlKODDFISCategory.SelectedItem.Text
    '        ddlCategory.Visible = False

    '        txtSubCategory.Visible = True
    '        txtSubCategory.Text = ddlKODDFISSubCategory.SelectedItem.Text
    '        ddlSubCategory.Visible = False

    '        pbIsAddingCategory = True
    '        pbIsAddingSubCategory = True

    '        Call ShowNewProduct()
    '        Call SetHelpStatus()
    '    End If
    'End Sub
    
    'Protected Function GetKODDFISDivisionCode() As String
    '    GetKODDFISDivisionCode = String.Empty
    '    Dim oDataReader As SqlDataReader
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim sSQL As String = "SELECT * FROM ClientData_KODDFIS_Divisions WHERE Category = '" & ddlKODDFISCategory.SelectedItem.Text & "'"
    '    Dim oCmd As New SqlCommand(sSQL, oConn)
    '    Try
    '        oConn.Open()
    '        oDataReader = oCmd.ExecuteReader()
    '        If oDataReader.HasRows Then
    '            oDataReader.Read()
    '            GetKODDFISDivisionCode = oDataReader("Division")
    '        End If
    '    Catch ex As SqlException
    '        WebMsgBox.Show("Error in GetKODDFISDivisionCode: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Function
    
    'Protected Function GetNextKODDFISCode(ByVal sDivision As String) As String
    '    Dim sCode As String
    '    Dim nHighestNumber As Integer = 0
    '    GetNextKODDFISCode = "0"
    '    Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE ProductCode LIKE '" & sDivision & "%' AND CustomerKey = " & Session("CustomerKey")
    '    Dim oDataReader As SqlDataReader
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim oCmd As New SqlCommand(sSQL, oConn)
    '    Try
    '        oConn.Open()
    '        oDataReader = oCmd.ExecuteReader()
    '        While oDataReader.Read()
    '            sCode = oDataReader("ProductCode")
    '            Dim sNumericPart As String = sCode.Substring(sDivision.Length, sCode.Length - sDivision.Length)
    '            If IsNumeric(sNumericPart) Then
    '                If CInt(sNumericPart) > nHighestNumber Then
    '                    nHighestNumber = CInt(sNumericPart)
    '                End If
    '            End If
    '        End While
    '        GetNextKODDFISCode = (nHighestNumber + 1).ToString("000")
    '    Catch ex As SqlException
    '        WebMsgBox.Show("Error in GetNextKODDFISCode: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    'End Function
    
    'Protected Sub btnBackFromNewKODDFISProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call BackToProductListPanel()
    'End Sub
    
    'Protected Sub cbKODDFISUseCategories_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim cb As CheckBox = sender
    '    If Not cb.Checked Then
    '        Call ShowNewProduct()
    '        Call SetHelpStatus()
    '    End If
    'End Sub

    Protected Sub btnBackFromConfigureCustomLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BackFromConfigureCustomLetter()
    End Sub
    
    Protected Sub BackFromConfigureCustomLetter()
        If pbIsAddingNew Then
            If bSetExplicitProductPermissionsFlag() Then
                Call ShowProductUserProfile()
            Else
                Call BackToProductListPanel()
            End If
        Else
            Call ShowProductDetail()
        End If
    End Sub
    
    Protected Sub SaveCustomLetterConfiguration()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CustomLetter_TemplateSet", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = plProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)

        Dim paramDocumentTemplate As SqlParameter = New SqlParameter("@DocumentTemplate", SqlDbType.NVarChar, 2000)
        paramDocumentTemplate.Value = fckedCustomLetterTemplate.Value
        oCmd.Parameters.Add(paramDocumentTemplate)

        Dim paramDocumentInstructions As SqlParameter = New SqlParameter("@DocumentInstructions", SqlDbType.NVarChar, 2000)
        paramDocumentInstructions.Value = tbCustomLetterInstructions.Text
        oCmd.Parameters.Add(paramDocumentInstructions)

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Error in SaveCustomLetterConfiguration: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetCustomLetterConfiguration()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_CustomLetter_TemplateGet", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int)
        oParam.Value = plProductKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                fckedCustomLetterTemplate.Value = oDataReader("DocumentTemplate")
                tbCustomLetterInstructions.Text = oDataReader("DocumentInstructions")
            Else
                fckedCustomLetterTemplate.Value = String.Empty
                tbCustomLetterInstructions.Text = String.Empty
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Error in GetCustomLetterConfiguration: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btnSaveCustomLetterConfiguration_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveCustomLetterConfiguration()
        Call BackFromConfigureCustomLetter()
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

    Protected Sub ddlItemsPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        dg_ProductList.PageSize = ddl.SelectedValue
        dg_ProductList.CurrentPageIndex = 0
        Call BindProductGridDispatcher(nCategoryMode:=pnCategoryMode)
    End Sub
    
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

    Protected Sub gvAuthoriseOrder_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvAuthoriseOrder.PageIndex = e.NewPageIndex
        Call GetPendingOrderAuthorisations()
    End Sub
    
    Protected Sub cbChangeAuthoriserOnAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            cbChangeAuthoriserOnAllSelective.Enabled = False
            cbChangeAuthoriserOnAllSelective.Checked = False
        Else
            cbChangeAuthoriserOnAllSelective.Enabled = True
        End If
    End Sub
    
    Protected Sub ddlUsersPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        grid_ProductUsers.PageSize = ddl.SelectedValue
        grid_ProductUsers.CurrentPageIndex = 0
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
    End Sub
    
    Protected Sub grid_ProductUsers_PageChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        grid_ProductUsers.CurrentPageIndex = e.NewPageIndex
        Call BindProductUserProfileGrid(txtProductUserSearch.Text, psSortValue)
    End Sub
    
    Protected Sub btnManagePublicationOwnerCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Function IsJupiterPODProduct(ByVal DataItem As Object) As Boolean
        IsJupiterPODProduct = False
        If IsJupiter() Then
            If Not IsDBNull(DataBinder.Eval(DataItem, "Misc2")) Then
                If IsNumeric(DataBinder.Eval(DataItem, "Misc2")) Then
                    If CInt(DataBinder.Eval(DataItem, "Misc2")) > 0 Then
                        IsJupiterPODProduct = True
                    End If
                End If
            End If
        End If
    End Function

    Protected Function GetJupiterPODProductType(ByVal DataItem As Object) As String
        GetJupiterPODProductType = String.Empty
        If Not IsDBNull(DataBinder.Eval(DataItem, "Misc2")) Then
            If IsNumeric(DataBinder.Eval(DataItem, "Misc2")) Then
                Dim sSQL As String = "SELECT PrintType FROM ClientData_Jupiter_PrintCost WHERE [id] = " & CInt(DataBinder.Eval(DataItem, "Misc2"))
                Dim dt = ExecuteQueryToDataTable(sSQL)
                If dt.Rows.Count > 0 Then
                    GetJupiterPODProductType = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
                End If
            End If
        End If
    End Function

    Property psSortValue() As String
        Get
            Dim o As Object = ViewState("PM_SortValue")
            If o Is Nothing Then
                Return "UserId"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_SortValue") = Value
            If Value.ToLower = "userid" Then
                lblSortValue.Text = "User ID"
            ElseIf Value.ToLower = "firstname" Then
                lblSortValue.Text = "First name"
            ElseIf Value.ToLower = "lastname" Then
                lblSortValue.Text = "Last name"
            End If
        End Set
    End Property

    Property plProductKey() As Long
        Get
            Dim o As Object = ViewState("PM_ProductKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PM_ProductKey") = Value
        End Set
    End Property
  
    Property psProdImageFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdImageFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdImageFolder") = Value
        End Set
    End Property
  
    Property psVirtualJPGFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualJPGFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualJPGFolder") = Value
        End Set
    End Property
  
    Property psProdThumbFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdThumbFolder") = Value
        End Set
    End Property
  
    Property psVirtualThumbFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualThumbFolder") = Value
        End Set
    End Property
  
    Property psProdPDFFolder() As String
        Get
            Dim o As Object = ViewState("PM_ProdPDFFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_ProdPDFFolder") = Value
        End Set
    End Property
  
    Property psVirtualPDFFolder() As String
        Get
            Dim o As Object = ViewState("PM_VirtualPDFFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_VirtualPDFFolder") = Value
        End Set
    End Property
  
    Property pbIsAddingNew() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingNew")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingNew") = Value
            If Value = True Then
                ResetForm()
                lblProductQuantity.Text = "0"
                txtProductCode.Enabled = True
                txtProductDate.Enabled = True
            Else
                txtProductCode.Enabled = False
                txtProductDate.Enabled = False
            End If
        End Set
    End Property

    Property pbIsAddingCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingCategory") = Value
            If Value = True Then
                txtCategory.Visible = True
                ddlCategory.Visible = False
                txtCategory.Focus()
            Else
                txtCategory.Visible = False
                ddlCategory.Visible = True
            End If
        End Set
    End Property

    Property pbIsAddingSubCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingSubCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingSubCategory") = Value
            If Value = True Then
                txtSubCategory.Visible = True
                ddlSubCategory.Visible = False
                txtSubCategory.Focus()
            Else
                txtSubCategory.Visible = False
                ddlSubCategory.Visible = True
            End If
        End Set
    End Property

    Property pbIsAddingSubSubCategory() As Boolean
        Get
            Dim o As Object = ViewState("PM_IsAddingSubSubCategory")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_IsAddingSubSubCategory") = Value
            If Value = True Then
                tbSubSubCategory.Visible = True
                ddlSubSubCategory.Visible = False
                tbSubSubCategory.Focus()
            Else
                tbSubSubCategory.Visible = False
                ddlSubSubCategory.Visible = True
            End If
        End Set
    End Property

    Property plPerCustomerConfiguration() As Long
        Get
            Dim o As Object = ViewState("PM_PerCustomerConfiguration")
            If o Is Nothing Then
                Return PER_CUSTOMER_CONFIGURATION_NONE
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PM_PerCustomerConfiguration") = Value
        End Set
    End Property
  
    Property plOwnerGroup() As Long
        Get
            Dim o As Object = ViewState("PM_OwnerGroup")
            If o Is Nothing Then
                Return PER_USERTYPE_OWNER_GROUP_NONE
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PM_OwnerGroup") = Value
        End Set
    End Property
  
    Property pnCategoryMode() As Integer
        Get
            Dim o As Object = ViewState("PM_CategoryMode")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("PM_CategoryMode") = Value
        End Set
    End Property
  
    Property pbUsesCategories() As Boolean
        Get
            Dim o As Object = ViewState("PM_UsesCategories")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_UsesCategories") = Value
        End Set
    End Property
  
    Property psCategory() As String
        Get
            Dim o As Object = ViewState("PM_Category")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_Category") = Value
        End Set
    End Property
  
    Property psSubCategory() As String
        Get
            Dim o As Object = ViewState("PM_SubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_SubCategory") = Value
        End Set
    End Property
  
    Property psSubSubCategory() As String
        Get
            Dim o As Object = ViewState("PM_SubSubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_SubSubCategory") = Value
        End Set
    End Property
  
    Property psDisplayMode() As String
        Get
            Dim o As Object = ViewState("PM_DisplayMode")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PM_DisplayMode") = Value
        End Set
    End Property
  
    Property pbOrderAuthorisation() As Boolean
        Get
            Dim o As Object = ViewState("PM_OrderAuthorisation")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_OrderAuthorisation") = Value
        End Set
    End Property

    Property pbProductAuthorisation() As Boolean
        Get
            Dim o As Object = ViewState("PM_ProductAuthorisation")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_ProductAuthorisation") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("PM_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_ProductOwners") = Value
        End Set
    End Property
   
    Property pbCalendarManagement() As Boolean
        Get
            Dim o As Object = ViewState("PM_CalendarManagement")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_CalendarManagement") = Value
        End Set
    End Property
   
    Property pbCustomLetters() As Boolean
        Get
            Dim o As Object = ViewState("PM_CustomLetters")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_CustomLetters") = Value
        End Set
    End Property
   
    Property pbSellingPrice() As Boolean
        Get
            Dim o As Object = ViewState("PM_SellingPrice")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_SellingPrice") = Value
        End Set
    End Property

    Property pbProductCredits() As Boolean
        Get
            Dim o As Object = ViewState("PM_ProductCredits")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PM_ProductCredits") = Value
        End Set
    End Property
        
    Protected Sub lnkbtnConfigureProductCredits_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        lblConfigureProductCreditsProductCode.Text = "for " & txtProductCode.Text
        Call PopulateProductCreditAvailableGroupsDropdown()
        pnlConfigureProductCredits.Visible = True
    End Sub

    Protected Sub PopulateProductCreditAvailableGroupsDropdown()
        Dim sSQL As String
        Dim dtUserGroups As DataTable
        'sSQL = "SELECT upg.GroupName, pcc.[id], ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), NextRefreshDateTime, 106), ' ', '-') AS varchar(20)),'(never)') 'NextRefreshDateTime', 'Every ' + CAST(RefreshInterval AS varchar(10)) + CASE(RefreshDaysOrMonths) WHEN 'M' THEN ' Month' ELSE ' Day' END + CASE(RefreshInterval) WHEN 1 THEN '' ELSE 's' END 'RefreshMessage', RefreshDaysOrMonths, RefreshInterval, CarryOverCredit FROM UP_UserPermissionGroups upg LEFT OUTER JOIN ProductCreditControl pcc ON upg.[id] = pcc.UserOrUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND pcc.LogisticProductKey = " & plProductKey & " ORDER BY GroupName"
        Call BindProductCreditControl()
        sSQL = "SELECT upg.[id], GroupName FROM UP_UserPermissionGroups upg WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND NOT [id] IN (SELECT upg.[id] FROM UP_UserPermissionGroups upg LEFT OUTER JOIN ProductCreditControl pcc ON upg.[id] = pcc.UserOrUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND pcc.LogisticProductKey = " & plProductKey & ") ORDER BY GroupName"
        dtUserGroups = ExecuteQueryToDataTable(sSQL)
        ddlProductCreditAvailableGroups.Items.Clear()
        ddlProductCreditAvailableGroups.Items.Add(New ListItem("- please select -", 0))
        ddlProductCreditAvailableGroups.Items.Add(New ListItem("- all groups -", 999999))
        For Each dr As DataRow In dtUserGroups.Rows
            ddlProductCreditAvailableGroups.Items.Add(New ListItem(dr("GroupName"), dr("id")))
        Next
    End Sub
    
    Protected Sub BindProductCreditControl()
        Dim sSQL As String
        Dim dtUserGroups As DataTable
        sSQL = "SELECT upg.GroupName, pcc.[id], Credit, EnforceCreditLimit, ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), NextRefreshDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), NextRefreshDateTime, 108)),1,5) AS varchar(20)),'(never)') 'NextRefreshDateTime', 'Every ' + CAST(RefreshInterval AS varchar(10)) + CASE(RefreshDaysOrMonths) WHEN 'M' THEN ' Month' ELSE ' Day' END + CASE(RefreshInterval) WHEN 1 THEN '' ELSE 's' END 'RefreshMessage', RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits FROM UP_UserPermissionGroups upg LEFT OUTER JOIN ProductCreditControl pcc ON upg.[id] = pcc.UserOrUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND pcc.LogisticProductKey = " & plProductKey & " ORDER BY GroupName"
        dtUserGroups = ExecuteQueryToDataTable(sSQL)
        gvProductCreditsIncludedGroups.DataSource = dtUserGroups
        gvProductCreditsIncludedGroups.DataBind()
    End Sub
    
    Protected Sub cbProductCredits_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            WebMsgBox.Show("After selecting the Product Credits option, please save your changes, then re-select this product.\r\nClick on the 'configure the product credits' link to finish configuring product credits.")
        Else
            WebMsgBox.Show("Warning: if you save this change you will lose any product credit configuration data for this product.")
        End If
    End Sub
    
    Protected Sub btnSaveProductCreditsConfiguration_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub

    Protected Sub btnCancelProductCreditsConfiguration_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductDetail()
    End Sub
    
    Protected Sub ddlProductCreditAvailableGroups_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        trEditConfigureProductCredits.Visible = False
        Dim ddl As DropDownList = sender
        Dim sSQL As String
        If ddl.SelectedValue = 999999 Then
            sSQL = "INSERT INTO ProductCreditControl (LogisticProductKey, Credit, EnforceCreditLimit, UserOrUserGroup, NextRefreshDateTime, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits) SELECT " & plProductKey & ", 0, 1, upg.[id], GETDATE(), 'M', 1, 'N', 0  FROM UP_UserPermissionGroups upg WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND NOT [id] IN (SELECT upg.[id] FROM UP_UserPermissionGroups upg LEFT OUTER JOIN ProductCreditControl pcc ON upg.[id] = pcc.UserOrUserGroup WHERE upg.CustomerKey = " & Session("CustomerKey") & " AND pcc.LogisticProductKey = " & plProductKey & ")"
            Call ExecuteQueryToDataTable(sSQL)
            Call PopulateProductCreditAvailableGroupsDropdown()
        ElseIf ddl.SelectedValue > 0 Then
            sSQL = "INSERT INTO ProductCreditControl (LogisticProductKey, Credit, EnforceCreditLimit, UserOrUserGroup, NextRefreshDateTime, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits) VALUES (" & plProductKey & ", 0, 1, " & ddl.SelectedValue & ", GETDATE(), 'M', 1, 'N', 0)"
            Call ExecuteQueryToDataTable(sSQL)
            Call PopulateProductCreditAvailableGroupsDropdown()
        End If
    End Sub
    
    Protected Sub lnkbtnProductCreditsEditEntry_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nID As Int32 = lnkbtn.CommandArgument
        trEditConfigureProductCredits.Visible = True
        Dim dr As DataRow = ExecuteQueryToDataTable("SELECT GroupName, Credit, EnforceCreditLimit, NextRefreshDateTime, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] WHERE pcc.[id] = " & nID).Rows(0)
        rntbCredit.Text = dr("Credit").ToString
        If dr("EnforceCreditLimit") = CREDIT_LIMIT_ENFORCE_TRUE Then
            cbConfigureProductCreditsEnforce.Checked = True
        Else
            cbConfigureProductCreditsEnforce.Checked = False
        End If
        'rdtpNextRefresh.SelectedDate = dr("NextRefreshDateTime")
        rntbRefreshInterval.Text = dr("RefreshInterval").ToString
        If dr("RefreshDaysOrMonths").ToString.ToLower = "d" Then
            rbConfigureProductCreditsIntervalDays.Checked = True
        Else
            rbConfigureProductCreditsIntervalMonths.Checked = True
        End If
        'If dr("CarryOverCredit") = "Y" Then
        '    cbConfigureProductCreditsCarryOverCredit.Checked = True
        'Else
        '    cbConfigureProductCreditsCarryOverCredit.Checked = False
        'End If
        lblUserGroup.Text = dr("GroupName")
        'rntbMaxCredits.Text = dr("MaxCredits").ToString
        btnConfigureProductCreditsSave.CommandArgument = nID
    End Sub

    Protected Sub lnkbtnProductCreditsRemoveEntry_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trEditConfigureProductCredits.Visible = False
        Dim lnkbtn As LinkButton = sender
        Dim nID As Int32 = lnkbtn.CommandArgument
        Dim nUserOrUserGroup As Int32 = ExecuteQueryToDataTable("SELECT UserOrUserGroup FROM ProductCreditControl WHERE [id] = " & nID).Rows(0).Item(0)
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE LogisticProductKey = " & plProductKey & " AND UserKey IN (SELECT [key] FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND ISNULL(UserGroup, 0) = " & nUserOrUserGroup & ")")
        Call ExecuteQueryToDataTable("DELETE FROM ProductCreditControl WHERE [id] = " & nID)
        Call PopulateProductCreditAvailableGroupsDropdown()
        Call BindProductCreditControl()
    End Sub
    
    Protected Sub btnConfigureProductCreditsSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        Dim nID As Int32 = btn.CommandArgument
        Dim sRefreshDaysOrMonths As String
        If rbConfigureProductCreditsIntervalDays.Checked Then
            sRefreshDaysOrMonths = "D"
        Else
            sRefreshDaysOrMonths = "M"
        End If
        Dim sCarryOverCredit As String
        'If cbConfigureProductCreditsCarryOverCredit.Checked Then
        '    sCarryOverCredit = "Y"
        'Else
        '    sCarryOverCredit = "N"
        'End If
        sCarryOverCredit = "N"
        rntbRefreshInterval.Text = rntbRefreshInterval.Text.Trim
        If rntbRefreshInterval.Text = String.Empty Then
            rntbRefreshInterval.Text = 0
        End If
        Dim nEnforceCreditLimit As Int32
        If cbConfigureProductCreditsEnforce.Checked Then
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_TRUE
        Else
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_FALSE
        End If
        'Dim sMaxCredits As String = rntbMaxCredits.Text
        Dim sMaxCredits As String = "0"
        'If sMaxCredits = String.Empty Then
        'sMaxCredits = 0
        'End If
        'Dim sSQL As String = "UPDATE ProductCreditControl SET Credit = " & rntbCredit.Text & ", NextRefreshDateTime = '" & Date.Parse(rdtpNextRefresh.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "', RefreshDaysOrMonths = '" & sRefreshDaysOrMonths & "', RefreshInterval = " & rntbRefreshInterval.Text & ", EnforceCreditLimit = " & nEnforceCreditLimit & ", CarryOverCredit = '" & sCarryOverCredit & "', MaxCredits = " & sMaxCredits & " WHERE [id] = " & nID
        Dim sSQL As String = "UPDATE ProductCreditControl SET Credit = " & rntbCredit.Text & ", NextRefreshDateTime = GETDATE(), RefreshDaysOrMonths = '" & sRefreshDaysOrMonths & "', RefreshInterval = " & rntbRefreshInterval.Text & ", EnforceCreditLimit = " & nEnforceCreditLimit & ", CarryOverCredit = '" & sCarryOverCredit & "', MaxCredits = " & sMaxCredits & " WHERE [id] = " & nID
        Call ExecuteQueryToDataTable(sSQL)
        Call BindProductCreditControl()
        trEditConfigureProductCredits.Visible = False
        Call SetTemplateChangeFlag()
    End Sub
    
    Protected Sub btnConfigureProductCreditsSaveAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        'Dim nID As Int32 = btn.CommandArgument
        Dim sRefreshDaysOrMonths As String
        If rbConfigureProductCreditsIntervalDays.Checked Then
            sRefreshDaysOrMonths = "D"
        Else
            sRefreshDaysOrMonths = "M"
        End If
        Dim sCarryOverCredit As String
        'If cbConfigureProductCreditsCarryOverCredit.Checked Then
        '    sCarryOverCredit = "Y"
        'Else
        '    sCarryOverCredit = "N"
        'End If
        sCarryOverCredit = "N"
        rntbRefreshInterval.Text = rntbRefreshInterval.Text.Trim
        If rntbRefreshInterval.Text = String.Empty Then
            rntbRefreshInterval.Text = 0
        End If
        Dim nEnforceCreditLimit As Int32
        If cbConfigureProductCreditsEnforce.Checked Then
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_TRUE
        Else
            nEnforceCreditLimit = CREDIT_LIMIT_ENFORCE_FALSE
        End If
        'Dim sMaxCredits As String = rntbMaxCredits.Text
        Dim sMaxCredits As String = "0"
        'If sMaxCredits = String.Empty Then
        'sMaxCredits = 0
        'End If
        'Dim sSQL As String = "UPDATE ProductCreditControl SET Credit = " & rntbCredit.Text & ", NextRefreshDateTime = '" & Date.Parse(rdtpNextRefresh.SelectedDate).ToString("dd-MMM-yyyy hh:mm:ss") & "', RefreshDaysOrMonths = '" & sRefreshDaysOrMonths & "', RefreshInterval = " & rntbRefreshInterval.Text & ", EnforceCreditLimit = " & nEnforceCreditLimit & ", CarryOverCredit = '" & sCarryOverCredit & "', MaxCredits = " & sMaxCredits & " WHERE LogisticProductKey = " & plProductKey
        Dim sSQL As String = "UPDATE ProductCreditControl SET Credit = " & rntbCredit.Text & ", NextRefreshDateTime = GETDATE(), RefreshDaysOrMonths = '" & sRefreshDaysOrMonths & "', RefreshInterval = " & rntbRefreshInterval.Text & ", EnforceCreditLimit = " & nEnforceCreditLimit & ", CarryOverCredit = '" & sCarryOverCredit & "', MaxCredits = " & sMaxCredits & " WHERE LogisticProductKey = " & plProductKey
        Call ExecuteQueryToDataTable(sSQL)
        Call BindProductCreditControl()
        trEditConfigureProductCredits.Visible = False
        Call SetTemplateChangeFlag()
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

    Protected Sub btnUsersReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub

    Protected Sub btnProductCredits_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlProductCreditsControl.Visible = True
    End Sub
    
    Protected Sub btnProductCreditsRefreshCredits_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call RemoveExpiredOrFutureCredits()
        'Dim sMessage As String = RefreshCredits()
        Call SetTemplateChangeFlag()
        'WebMsgBox.Show("Credit refresh completed. " & sMessage)
        WebMsgBox.Show("Product credits are being refreshed.")
    End Sub

    Protected Function GetRemainingCredit(ByVal nUserKey As Int32) As Int32
        GetRemainingCredit = 0
    End Function
    
    Protected Function GetProductCodeDescriptionFromKey(ByVal nLogisticProductKey As Int32) As String
        GetProductCodeDescriptionFromKey = ExecuteQueryToDataTable("SELECT ProductCode + ' ' + ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey).Rows(0).Item(0)
    End Function
    
    Protected Sub RemoveExpiredOrFutureCredits()
        Dim sSQL As String
        sSQL = "DELETE FROM ProductCredits WHERE (CreditStartTime > GETDATE() OR CreditEndTime < GETDATE()) AND UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Function RefreshCredits() As String
        Dim sSQL As String = "SELECT pcc.UserOrUserGroup, pcc.LogisticProductKey, pcc.Credit, pcc.EnforceCreditLimit, ISNULL(pcc.NextRefreshDateTime, '1-Jan-2013') 'NextRefreshDateTime', pcc.RefreshDaysOrMonths, pcc.RefreshInterval, ISNULL(pcc.CarryOverCredit, 'N') 'CarryOverCredit', ISNULL(pcc.MaxCredits, 0) 'MaxCredits' FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] WHERE upg.CustomerKey = " & Session("CustomerKey")
        ' can remove NextRefreshTime which is no longer used
        Dim dtUserGroups As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim nTotalChanges As Int32 = 0
        For Each drUserGroup As DataRow In dtUserGroups.Rows
            Dim nUserOrUserGroup As Int32 = drUserGroup("UserOrUserGroup")
            Dim nLogisticProductKey As Int32 = drUserGroup("LogisticProductKey")
            Dim nCredit As Int32 = drUserGroup("Credit")
            Dim nEnforceCreditLimit As Int32 = drUserGroup("EnforceCreditLimit")
            'Dim dateNextRefreshDateTime As DateTime = drUserGroup("NextRefreshDateTime")
            Dim dateNextRefreshDateTime As DateTime = DateTime.Now
            Dim sRefreshDaysOrMonths As String = drUserGroup("RefreshDaysOrMonths")
            Dim nRefreshInterval As Int32 = drUserGroup("RefreshInterval")
            'Dim sCarryOverCredit As String = drUserGroup("CarryOverCredit")
            'Dim nMaxCredits As Int32 = drUserGroup("MaxCredits")
            
            sSQL = "SELECT [key] FROM UserProfile up WHERE UserGroup = " & nUserOrUserGroup & " AND up.Status = 'Active' AND DeletedFlag = 0"

            Dim sEndDateExpression As String
            'If sRefreshDaysOrMonths.ToLower = "d" Then
            '    sEndDateExpression = "DATEADD(DAY, " & nRefreshInterval & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "')"
            'Else
            '    sEndDateExpression = "DATEADD(MONTH, " & nRefreshInterval & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "')"
            'End If
            If sRefreshDaysOrMonths.ToLower = "d" Then
                sEndDateExpression = "DATEADD(DAY, " & nRefreshInterval & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "')"
            Else
                sEndDateExpression = "DATEADD(MONTH, " & nRefreshInterval & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "')"
            End If

            Dim dtUsersInUserGroup As DataTable = ExecuteQueryToDataTable(sSQL)
            For Each drUser As DataRow In dtUsersInUserGroup.Rows
                Dim nUserKey As Int32 = drUser("key")
                sSQL = "IF NOT EXISTS(SELECT 1 FROM ProductCredits WHERE UserKey = " & nUserKey & " AND LogisticProductKey = " & nLogisticProductKey & ") "
                sSQL &= "INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) VALUES ("
                sSQL &= nLogisticProductKey & ", " & nUserKey & ", " & nCredit & ", " & nCredit & ", " & nEnforceCreditLimit & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "', " & sEndDateExpression & ")"
                'If nRefreshInterval > 0 And dateNextRefreshDateTime <= Now Then
                'Dim nUserKey As Int32 = drUser("key")
                'Dim nProductCredit As Int32 = 0
                'If sCarryOverCredit = "Y" Then
                'nProductCredit = GetRemainingCredit(drUser("key"))
                'End If
                'sSQL = "DELETE FROM ProductCredits WHERE UserKey = " & nUserKey & " AND LogisticProductKey = " & nLogisticProductKey
                'Call ExecuteQueryToDataTable(sSQL)
                'nProductCredit += nCredit
                'If nMaxCredits > 0 Then
                '    If nProductCredit > nMaxCredits Then
                '        nProductCredit = nMaxCredits
                '    End If
                'End If
                'sSQL = "INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) VALUES ("
                ''sSQL &= nLogisticProductKey & ", " & nUserKey & ", " & nProductCredit & ", " & nProductCredit & ", " & nEnforceCreditLimit & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "', " & sEndDateExpression & ")"
                'sSQL &= nLogisticProductKey & ", " & nUserKey & ", " & nProductCredit & ", " & nProductCredit & ", " & nEnforceCreditLimit & ", '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm") & "', " & sEndDateExpression & ")"
                Call ExecuteQueryToDataTable(sSQL)
                nTotalChanges += 1
                'End If
            Next
            'sSQL = "UPDATE ProductCreditControl SET NextRefreshDateTime = " & sEndDateExpression & " WHERE UserOrUserGroup = " & nUserOrUserGroup
            sSQL = "UPDATE ProductCreditControl SET NextRefreshDateTime = '" & DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss") & "' WHERE UserOrUserGroup = " & nUserOrUserGroup
            Call ExecuteQueryToDataTable(sSQL)
        Next
        RefreshCredits = "Total users refreshed: " & nTotalChanges
        'WebMsgBox.Show("Total users refreshed: " & nTotalChanges)
    End Function
    
    Protected Function Bold(ByVal sString As String) As String
        Bold = "<b>" & sString & "</b>"
    End Function
   
    Protected Sub btnProductCreditsTemplateReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dt As DataTable
        Dim sSQL As String = String.Empty
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Product Credits - Templates Report")
        sbText.Append(Bold("PRODUCT CREDITS - TEMPLATES REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("This report is divided into sections. <b>Section 1</b> shows the products for which credit control is enabled. <b>Section 2</b> shows, for each product, the template for each user group. <b>Section 3</b> lists the user groups that have one or more products with a product credit template defined. <b>Section 4</b> shows, for each user group, the products with a product credit template defined.")
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("<hr />")
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT pcc.LogisticProductKey, lp.ProductCode + ' ' + lp.ProductDescription 'Product' FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] INNER JOIN LogisticProduct lp ON pcc.LogisticProductKey = lp.LogisticProductKey WHERE upg.CustomerKey =  " & Session("CustomerKey") & " ORDER BY pcc.LogisticProductKey", "Product", "LogisticProductKey")
        sbText.Append(Bold("1. Products under credit control (" & oListItemCollection.Count & ") are:"))
        Call NewLine(sbText)
        If oListItemCollection.Count > 0 Then
            For Each liProduct As ListItem In oListItemCollection
                sbText.Append(liProduct.Text)
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("(none)")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
        'Dim oListItemCollection5 As ListItemCollection = ExecuteQueryToListItemCollection("SELECT [Key] UserKey, FirstName + ' ' + LastName + ' (' + UserId + ')' UserName FROM UserProfile WHERE Type = 'User' AND CustomerKey = " & Session("CustomerKey") & " AND NOT ISNULL(UserGroup,0) IN (SELECT [id] FROM UP_UserPermissionGroups WHERE CustomerKey = " & Session("CustomerKey") & ") ORDER BY FirstName", "UserKey", "UserName")
        sbText.Append("<hr />")
        sbText.Append(Bold("2. User Group settings by Product"))
        Call NewLine(sbText)
        If oListItemCollection.Count > 0 Then
            For Each liProduct As ListItem In oListItemCollection
                Dim nLogisticProductKey As Int32 = liProduct.Value
                sbText.Append("PRODUCT: ")
                sbText.Append(Bold(GetProductCodeDescriptionFromKey(nLogisticProductKey)))
                Call NewLine(sbText)
                sSQL = "SELECT ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), NextRefreshDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), NextRefreshDateTime, 108)),1,5) AS varchar(20)),'(never)') 'NextRefreshDateTime', Credit, EnforceCreditLimit, 'Every ' + CAST(RefreshInterval AS varchar(10)) + CASE(RefreshDaysOrMonths) WHEN 'M' THEN ' Month' ELSE ' Day' END + CASE(RefreshInterval) WHEN 1 THEN '' ELSE 's' END 'RefreshMessage', CarryOverCredit, ISNULL(MaxCredits, 0), upg.GroupName, lp.ProductCode + ' ' + lp.ProductDescription 'Product' FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] INNER JOIN LogisticProduct lp ON pcc.LogisticProductKey = lp.LogisticProductKey WHERE pcc.LogisticProductKey = " & nLogisticProductKey & " ORDER BY upg.GroupName"
                dt = ExecuteQueryToDataTable(sSQL)
                For Each dr As DataRow In dt.Rows
                    Dim s As String = String.Empty
                    sbText.Append("GROUP: ")
                    s = Bold(dr("GroupName").ToString.PadRight(25, Convert.ToChar("~")))
                    sbText.Append(s)
                    'sbText.Append(Bold(dr("GroupName")).ToString.PadRight(15, Convert.ToChar("~")))
                    sbText.Append(" CREDIT: ")
                    s = Bold(dr("Credit").ToString.PadRight(4, Convert.ToChar("~")))
                    sbText.Append(s)
                    'sbText.Append(Bold(dr("CreditAmount")).ToString.PadRight(4, Convert.ToChar("~")))
                    sbText.Append("~~~~~~~~~~~~NEXT REFRESH: ")
                    sbText.Append(Bold(dr("NextRefreshDateTime")))
                    sbText.Append("~~~~~~~~~~~~FREQUENCY: ")
                    sbText.Append(Bold(dr("RefreshMessage")))
                    Dim sCarryOverCredit As String = dr("CarryOverCredit")
                    If sCarryOverCredit = "y" Then
                        sbText.Append(" (credit carried over, ")
                        Dim nMaxCredits As Int32 = dr("MaxCredits")
                        If nMaxCredits = 0 Then
                            sbText.Append("no maximum)")
                        Else
                            sbText.Append("maximum " & nMaxCredits.ToString & ")")
                        End If
                    End If
                    Call NewLine(sbText)
                Next
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("(none)")
        End If

        Call NewLine(sbText)
        sbText.Append("<hr />")
        Call NewLine(sbText)

        Dim oListItemCollectionA As ListItemCollection = ExecuteQueryToListItemCollection("SELECT DISTINCT GroupName, UserOrUserGroup FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] INNER JOIN LogisticProduct lp ON pcc.LogisticProductKey = lp.LogisticProductKey WHERE upg.CustomerKey = " & Session("CustomerKey") & " ORDER BY GroupName", "GroupName", "UserOrUserGroup")
        sbText.Append(Bold("3. User groups with products under credit control (" & oListItemCollectionA.Count & ") are:"))
        Call NewLine(sbText)
        If oListItemCollectionA.Count > 0 Then
            For Each liGroup As ListItem In oListItemCollectionA
                sbText.Append(liGroup.Text)
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("(none)")
        End If
        Call NewLine(sbText)

        Call NewLine(sbText)
        sbText.Append("<hr />")
        Call NewLine(sbText)

        sbText.Append(Bold("4. Products by User Group"))
        Call NewLine(sbText)
        If oListItemCollectionA.Count > 0 Then
            For Each liGroup As ListItem In oListItemCollectionA
                Dim nUserOrUserGroup As Int32 = liGroup.Value
                sbText.Append("GROUP: ")
                sbText.Append(Bold(liGroup.Text))
                Call NewLine(sbText)
                sSQL = "SELECT ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), NextRefreshDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), NextRefreshDateTime, 108)),1,5) AS varchar(20)),'(never)') 'NextRefreshDateTime', Credit, EnforceCreditLimit, 'Every ' + CAST(RefreshInterval AS varchar(10)) + CASE(RefreshDaysOrMonths) WHEN 'M' THEN ' Month' ELSE ' Day' END + CASE(RefreshInterval) WHEN 1 THEN '' ELSE 's' END 'RefreshMessage', CarryOverCredit, upg.GroupName, lp.ProductCode + ' ' + lp.ProductDescription 'Product' FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] INNER JOIN LogisticProduct lp ON pcc.LogisticProductKey = lp.LogisticProductKey WHERE pcc.UserOrUserGroup = " & nUserOrUserGroup & " ORDER BY lp.ProductCode"
                dt = ExecuteQueryToDataTable(sSQL)
                For Each dr As DataRow In dt.Rows
                    Dim s As String = String.Empty
                    sbText.Append("PRODUCT: ")
                    s = Bold(dr("Product").ToString.PadRight(25, Convert.ToChar("~")))
                    sbText.Append(s)
                    'sbText.Append(Bold(dr("GroupName")).ToString.PadRight(15, Convert.ToChar("~")))
                    sbText.Append(" CREDIT: ")
                    s = Bold(dr("Credit").ToString.PadRight(4, Convert.ToChar("~")))
                    sbText.Append(s)

                    sbText.Append(" ENFORCE: ")
                    If dr("Credit") = 0 Then
                        sbText.Append(Bold("NO"))
                    Else
                        sbText.Append(Bold("YES"))
                    End If
                
                    'sbText.Append(Bold(dr("CreditAmount")).ToString.PadRight(4, Convert.ToChar("~")))
                    sbText.Append("~~~~~~~~~~~~NEXT REFRESH: ")
                    sbText.Append(Bold(dr("NextRefreshDateTime")))
                    sbText.Append("~~~~~~~~~~~~FREQUENCY: ")
                    sbText.Append(Bold(dr("RefreshMessage")))
                    Dim sCarryOverCredit As String = dr("CarryOverCredit")
                    If sCarryOverCredit = "y" Then
                        sbText.Append(" (credit carried over, ")
                        Dim nMaxCredits As Int32 = dr("MaxCredits")
                        If nMaxCredits = 0 Then
                            sbText.Append("no maximum)")
                        Else
                            sbText.Append("maximum " & nMaxCredits.ToString & ")")
                        End If
                    End If
                    Call NewLine(sbText)
                Next
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("(none)")
        End If

        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString.Replace("~", "&nbsp;"), "Product Credits Settings Report " & DateTime.Now.ToString("dd-MMM-yyyy @ hh:mm:ss"))
    End Sub

    Protected Sub btnResetCredits_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call ExecuteQueryToDataTable("UPDATE pcc SET pcc.NextRefreshDateTime = DATEADD(MINUTE, -2, GETDATE()) FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] WHERE upg.CustomerKey = " & Session("CustomerKey"))
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & ")")
        Call SetTemplateChangeFlag()
        WebMsgBox.Show("Product credits are being recreated.")
        'Dim sMessage As String = RefreshCredits()
        'WebMsgBox.Show("Credit reset completed. " & sMessage)
    End Sub
    
    Protected Sub lnkbtnProductCreditsApplyUserGroup_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trEditConfigureProductCredits.Visible = False

        Dim lnkbtn As LinkButton = sender
        Dim nID As Int32 = lnkbtn.CommandArgument
        Dim nUserOrUserGroup As Int32 = ExecuteQueryToDataTable("SELECT UserOrUserGroup FROM ProductCreditControl WHERE [id] = " & nID).Rows(0).Item(0)
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE LogisticProductKey = " & plProductKey & " AND UserKey IN (SELECT [key] FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND ISNULL(UserGroup, 0) = " & nUserOrUserGroup & ")")
        Call SetTemplateChangeFlag()
        Exit Sub
        
        'Dim drRefreshDefaults As DataRow = ExecuteQueryToDataTable("SELECT GroupName, CreditAmount, NextRefreshDateTime, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits FROM ProductCreditControl pcc INNER JOIN UP_UserPermissionGroups upg ON pcc.UserOrUserGroup = upg.[id] WHERE pcc.[id] = " & nID).Rows(0)
        Dim drRefreshDefaults As DataRow = ExecuteQueryToDataTable("SELECT Credit, EnforceCreditLimit, NextRefreshDateTime, RefreshDaysOrMonths, RefreshInterval, CarryOverCredit, MaxCredits FROM ProductCreditControl pcc WHERE pcc.[id] = " & nID).Rows(0)
        Dim nCredit As Int32 = drRefreshDefaults("Credit")
        Dim nEnforceCreditLimit As Int32 = drRefreshDefaults("EnforceCreditLimit")
        Dim dateNextRefreshDateTime As DateTime = drRefreshDefaults("NextRefreshDateTime")
        Dim sRefreshDaysOrMonths As String = drRefreshDefaults("RefreshDaysOrMonths")
        Dim nRefreshInterval As Int32 = drRefreshDefaults("RefreshInterval")
        Dim sCarryOverCredit As String = drRefreshDefaults("CarryOverCredit")
        'Dim nMaxCredits As Int32 = drRefreshDefaults("MaxCredits")

        Dim sEndDateExpression As String
        If sRefreshDaysOrMonths.ToLower = "d" Then
            sEndDateExpression = "DATEADD(DAY, " & nRefreshInterval & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "')"
        Else
            sEndDateExpression = "DATEADD(MONTH, " & nRefreshInterval & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "')"
        End If
        'Dim dt As DataTable
        'Dim sSQL As String = "INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) SELECT " & plProductKey & ", " & "[key], " & nCredit & ", " & nCredit & ", " & nEnforceCreditLimit & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "', " & sEndDateExpression & " FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND ISNULL(UserGroup, 0) = " & nUserOrUserGroup
        Call ExecuteQueryToDataTable("INSERT INTO ProductCredits (LogisticProductKey, UserKey, StartCredit, RemainingCredit, EnforceCreditLimit, CreditStartDateTime, CreditEndDateTime) SELECT " & plProductKey & ", " & "[key], " & nCredit & ", " & nCredit & ", " & nEnforceCreditLimit & ", '" & dateNextRefreshDateTime.ToString("dd-MMM-yyyy hh:mm") & "', " & sEndDateExpression & " FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND ISNULL(UserGroup, 0) = " & nUserOrUserGroup)
    End Sub
    
    Protected Sub AddQuantityToJupiterPODProduct(ByVal nLogisticProductKey As Int32)
        Dim sSQL As String
        sSQL = "SELECT LogisticProductQuantity FROM LogisticProductLocation WHERE LogisticProductKey = " & nLogisticProductKey & " AND WarehouseBayKey = " & DEMO_BAY_KEY
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDT.Rows.Count > 0 Then
            If oDT.Rows.Count = 1 Then
                sSQL = "UPDATE LogisticProductLocation SET LogisticProductQuantity = 999999 WHERE WarehouseBayKey = " & DEMO_BAY_KEY & " AND LogisticProductKey = " & nLogisticProductKey
            Else
                WebMsgBox.Show("Error - multiple instances of one product in a single location.")
            End If
        Else
            sSQL = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES (" & nLogisticProductKey & ", " & DEMO_BAY_KEY & ", 100000, GETDATE())"
        End If
        Call ExecuteQueryToDataTable(sSQL)
    End Sub

    Protected Sub PermissionJupiterPODProduct(ByVal nLogisticProductKey As Int32)
        Dim sSQL As String = "UPDATE UserProductProfile SET AbleToView = 1, AbleToPick = 1, AbleToEdit = 0, AbleToArchive = 0, AbleToDelete = 0, ApplyMaxGrab = 1, MaxGrabQty = 1000 WHERE ProductKey = " & nLogisticProductKey & " AND UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & CUSTOMER_JUPITER & ")"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub SetJupiterPODProductArchiveFlag(ByVal nLogisticProductKey As Int32, ByVal sArchiveFlag As String)
        Dim sSQL As String = "UPDATE LogisticProduct SET ArchiveFlag = '" & sArchiveFlag & "' WHERE LogisticProductKey = " & nLogisticProductKey
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    'Protected Sub ddlPrintType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim ddl As DropDownList = sender
    '    If ddl.SelectedIndex > 0 Then
    '        If pbIsAddingNew Then
    '            chkArchivedFlag.Checked = True
    '        End If
    '    End If
    'End Sub

    Protected Sub ddlPODPageCount_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            Dim nJupiterPODIndex As Int32 = GetJupiterPODIndex()
            If nJupiterPODIndex = 0 Then
                lblPODPrintType.Text = "(undefined)"
            Else
                lblPODPrintType.Text = ExecuteQueryToDataTable("SELECT PrintType FROM ClientData_Jupiter_PrintCost WHERE [id] = " & nJupiterPODIndex).Rows(0).Item(0)
            End If
            If pbIsAddingNew Then
                chkArchivedFlag.Checked = True
            End If
        Else
            lblPODPrintType.Text = "(undefined)"
        End If
    End Sub

    Protected Sub cbOnDemand_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        'ddlPrintType.SelectedIndex = 0
        If Not cb.Checked Then
            'rfdPrintType.Enabled = False
            'rfdPrintType.Visible = False
            'lblLegendPrintType.Visible = False
            'ddlPrintType.Visible = False
            Call SetOnDemandControlsVisibility(False)
        Else
            'rfdPrintType.Enabled = True
            'rfdPrintType.Visible = True
            'lblLegendPrintType.Visible = True
            'ddlPrintType.Visible = True
            txtProductDate.Text = txtProductDate.Text.Trim
            If txtProductDate.Text = String.Empty Then
                txtProductDate.Text = Date.Now.ToString("ddMMMyyyy")
            End If
            ddlPODPageCount.SelectedIndex = 0
            rbPOD120gsm.Checked = True
            rbPODSizeA4.Checked = True
            Call SetOnDemandControlsVisibility(True)
        End If
    End Sub
    
    Protected Sub SetOnDemandControlsVisibility(ByVal bVisible As Boolean)
        lblPODPrintType.Visible = bVisible
        rfdPODPageCount.Visible = bVisible
        lblLegendPODPageCount.Visible = bVisible
        'ddlPrintType.Visible = bVisible
        ddlPODPageCount.Visible = bVisible
        lblLegendPODSize.Visible = bVisible
        rbPODSizeA4.Visible = bVisible
        rbPODSizeA5.Visible = bVisible
        lblLegendPODStock.Visible = bVisible
        rbPOD120gsm.Visible = bVisible
        rbPOD150gsm.Visible = bVisible
        rbPOD200gsm.Visible = bVisible
    End Sub
                                                
    Protected Sub btnConfigureProductCreditsCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trEditConfigureProductCredits.Visible = False
        rntbCredit.Text = String.Empty
        cbConfigureProductCreditsEnforce.Checked = False
        rntbRefreshInterval.Text = String.Empty
        rbConfigureProductCreditsIntervalMonths.Checked = False
        rbConfigureProductCreditsIntervalDays.Checked = False
    End Sub
    
    Protected Function gvProductCreditType(ByVal DataItem As Object) As String
        If CBool(DataBinder.Eval(DataItem, "EnforceCreditLimit")) Then
            gvProductCreditType = "ENFORCED"
        Else
            gvProductCreditType = "OVERDRAFT"
        End If
    End Function

    Protected Sub btnRemoveAllTemplates_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        sSQL = "DELETE FROM ProductCreditControl WHERE LogisticProductKey IN (SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
        Call ExecuteQueryToDataTable("DELETE FROM ProductCredits WHERE UserKey IN (SELECT [key] FROM UserProfile WHERE CustomerKey = " & Session("CustomerKey") & ")")
        Call SetTemplateChangeFlag()
        WebMsgBox.Show("All product credit templates have been removed. No product credits are now in force.")
    End Sub
    
    Protected Sub btnRemoveAllTemplatesForThisProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        sSQL = "DELETE FROM ProductCreditControl WHERE LogisticProductKey = " & plProductKey
        Call ExecuteQueryToDataTable(sSQL)
        WebMsgBox.Show("All templates for this product have been removed.\n\nThere may still be product credits in force for the product. To remove them click 'refresh all credits for this product'.")
        'plProductKey
    End Sub

    Protected Sub btnRefreshAllCreditsForThisProduct_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        sSQL = "DELETE FROM ProductCredits WHERE LogisticProductKey = " & plProductKey
        Call ExecuteQueryToDataTable(sSQL)
        Call SetTemplateChangeFlag()
        WebMsgBox.Show("Product credits for this product are now being refreshed.")
    End Sub
    
    Protected Function SetUploadPDFVisibility() As Boolean
        Return False
    End Function
    
    Protected Sub rbPOD_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            If rb.ID.Contains("PODSize") Then
                rbPODSizeA4.Font.Bold = False
                rbPODSizeA5.Font.Bold = False
                rb.Font.Bold = True
            End If
            If rb.ID.Contains("gsm") Then
                rbPOD120gsm.Font.Bold = False
                rbPOD150gsm.Font.Bold = False
                rbPOD200gsm.Font.Bold = False
                rb.Font.Bold = True
            End If
        End If
        Dim nJupiterPODIndex As Int32 = GetJupiterPODIndex()
        If nJupiterPODIndex = 0 Then
            lblPODPrintType.Text = "(undefined)"
        Else
            lblPODPrintType.Text = ExecuteQueryToDataTable("SELECT PrintType FROM ClientData_Jupiter_PrintCost WHERE [id] = " & nJupiterPODIndex).Rows(0).Item(0)
        End If
    End Sub

    Protected Sub lnkbtnCategoriesReport_Click(sender As Object, e As System.EventArgs)
        Call CategoriesReport_Click()
    End Sub

    Protected Sub CategoriesReport_Click()
        Dim dt As DataTable
        Dim sSQL As String
        Dim sbText As New StringBuilder
        Call AddHTMLPreamble(sbText, "Categories Report")
        sbText.Append(Bold("CATEGORIES REPORT"))
        Call NewLine(sbText)
        sbText.Append("report generated " & DateTime.Now)
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append(Bold("LIST OF CATEGORIES"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT DISTINCT ProductCategory FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " AND ProductCategory <> '' ORDER BY ProductCategory"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                sbText.Append(dr(0))
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No categories defined.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
        sbText.Append(Bold("LIST OF SUB-CATEGORIES"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT DISTINCT ISNULL(SubCategory, '') 'SubCategory' FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " AND SubCategory <> '' ORDER BY SubCategory"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                sbText.Append(dr(0))
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No sub-categories defined.")
        End If

        If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
            sSQL = "SELECT DISTINCT ISNULL(SubCategory2, '') 'SubCategory2' FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " AND SubCategory2 <> '' ORDER BY SubCategory2"
            dt = ExecuteQueryToDataTable(sSQL)
            If dt.Rows.Count > 0 Then
                Call NewLine(sbText)
                Call NewLine(sbText)
                sbText.Append(Bold("LIST OF SUB-SUB-CATEGORIES"))
                Call NewLine(sbText)
                Call NewLine(sbText)
                For Each dr As DataRow In dt.Rows
                    sbText.Append(dr(0))
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("No sub-sub-categories defined.")
            End If
            Call NewLine(sbText)
            Call NewLine(sbText)
        End If
        
        sbText.Append(Bold("PRODUCTS WITH NO CATEGORY"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(ProductCategory, '') = '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                sbText.Append("Product code: ")
                sbText.Append(Bold(dr("ProductCode")))
                If dr("ProductCode").ToString.ToUpper = "Y" Then
                    sbText.Append(Bold("  (ARCHIVED)"))
                End If
                Call NewLine(sbText)
                If dr("ProductDate") = String.Empty Then
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold("(none)"))
                Else
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold(dr("ProductDate")))
                End If
                Call NewLine(sbText)
                sbText.Append("Product description: ")
                sbText.Append(Bold(dr("ProductDescription")))
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No products without a category.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)

        sbText.Append(Bold("PRODUCTS WITH NO SUBCATEGORY"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(SubCategory, '') = '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                sbText.Append("Product code: ")
                sbText.Append(Bold(dr("ProductCode")))
                If dr("ProductCode").ToString.ToUpper = "Y" Then
                    sbText.Append(Bold("  (ARCHIVED)"))
                End If
                Call NewLine(sbText)
                If dr("ProductDate") = String.Empty Then
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold("(none)"))
                Else
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold(dr("ProductDate")))
                End If
                Call NewLine(sbText)
                sbText.Append("Product description: ")
                sbText.Append(Bold(dr("ProductDescription")))
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No products without a subcategory.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)

        If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
            sbText.Append(Bold("PRODUCTS WITH NO SUB-SUBCATEGORY"))
            Call NewLine(sbText)
            Call NewLine(sbText)
            sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(SubCategory2, '') = '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
            dt = ExecuteQueryToDataTable(sSQL)
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    sbText.Append("Product code: ")
                    sbText.Append(Bold(dr("ProductCode")))
                    If dr("ProductCode").ToString.ToUpper = "Y" Then
                        sbText.Append(Bold("  (ARCHIVED)"))
                    End If
                    Call NewLine(sbText)
                    If dr("ProductDate") = String.Empty Then
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold("(none)"))
                    Else
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold(dr("ProductDate")))
                    End If
                    Call NewLine(sbText)
                    sbText.Append("Product description: ")
                    sbText.Append(Bold(dr("ProductDescription")))
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("No products without a sub-subcategory.")
            End If
            Call NewLine(sbText)
            Call NewLine(sbText)
        End If

        sbText.Append(Bold("PRODUCTS ASSIGNED TO CATEGORY"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT DISTINCT ISNULL(ProductCategory, '') 'ProductCategory' FROM LogisticProduct WHERE ProductCategory <> '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCategory"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                Dim sProductCategory As String = dr("ProductCategory")
                sbText.Append("Category: ")
                sbText.Append(Bold(sProductCategory))
                Call NewLine(sbText)
                sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(ProductCategory, '') = '" & sProductCategory.Replace("'", "''") & "' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
                Dim dt2 As DataTable = ExecuteQueryToDataTable(sSQL)
                For Each dr2 As DataRow In dt2.Rows
                    sbText.Append("Product code: ")
                    sbText.Append(Bold(dr2("ProductCode")))
                    If dr2("ProductCode").ToString.ToUpper = "Y" Then
                        sbText.Append(Bold("  (ARCHIVED)"))
                    End If
                    Call NewLine(sbText)
                    If dr2("ProductDate") = String.Empty Then
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold("(none)"))
                    Else
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold(dr2("ProductDate")))
                    End If
                    Call NewLine(sbText)
                    sbText.Append("Product description: ")
                    sbText.Append(Bold(dr2("ProductDescription")))
                    Call NewLine(sbText)
                Next
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No categories defined.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
       
        sbText.Append(Bold("PRODUCTS ASSIGNED TO SUBCATEGORY"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT DISTINCT ISNULL(SubCategory, '') 'SubCategory' FROM LogisticProduct WHERE SubCategory <> '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY SubCategory"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                Dim sSubCategory As String = dr("SubCategory")
                sbText.Append("Subcategory: ")
                sbText.Append(Bold(sSubCategory))
                Call NewLine(sbText)
                sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(SubCategory, '') = '" & sSubCategory.Replace("'", "''") & "' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
                Dim dt2 As DataTable = ExecuteQueryToDataTable(sSQL)
                For Each dr2 As DataRow In dt2.Rows
                    sbText.Append("Product code: ")
                    sbText.Append(Bold(dr2("ProductCode")))
                    If dr2("ProductCode").ToString.ToUpper = "Y" Then
                        sbText.Append(Bold("  (ARCHIVED)"))
                    End If
                    Call NewLine(sbText)
                    If dr2("ProductDate") = String.Empty Then
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold("(none)"))
                    Else
                        sbText.Append("Product value/date: ")
                        sbText.Append(Bold(dr2("ProductDate")))
                    End If
                    Call NewLine(sbText)
                    sbText.Append("Product description: ")
                    sbText.Append(Bold(dr2("ProductDescription")))
                    Call NewLine(sbText)
                Next
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No subcategories defined.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)
       
        If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
            sbText.Append(Bold("PRODUCTS ASSIGNED TO SUB-SUBCATEGORY"))
            Call NewLine(sbText)
            Call NewLine(sbText)
            sSQL = "SELECT DISTINCT ISNULL(SubCategory2, '') 'SubCategory2' FROM LogisticProduct WHERE SubCategory2 <> '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY SubCategory2"
            dt = ExecuteQueryToDataTable(sSQL)
            If dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim sSubCategory2 As String = dr("SubCategory2")
                    sbText.Append("Sub-Subcategory: ")
                    sbText.Append(Bold(sSubCategory2))
                    Call NewLine(sbText)
                    sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ArchiveFlag FROM LogisticProduct WHERE ISNULL(SubCategory2, '') = '" & sSubCategory2.Replace("'", "''") & "' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
                    Dim dt2 As DataTable = ExecuteQueryToDataTable(sSQL)
                    For Each dr2 As DataRow In dt2.Rows
                        sbText.Append("Product code: ")
                        sbText.Append(Bold(dr2("ProductCode")))
                        If dr2("ProductCode").ToString.ToUpper = "Y" Then
                            sbText.Append(Bold("  (ARCHIVED)"))
                        End If
                        Call NewLine(sbText)
                        If dr2("ProductDate") = String.Empty Then
                            sbText.Append("Product value/date: ")
                            sbText.Append(Bold("(none)"))
                        Else
                            sbText.Append("Product value/date: ")
                            sbText.Append(Bold(dr2("ProductDate")))
                        End If
                        Call NewLine(sbText)
                        sbText.Append("Product description: ")
                        sbText.Append(Bold(dr2("ProductDescription")))
                        Call NewLine(sbText)
                    Next
                    Call NewLine(sbText)
                Next
            Else
                sbText.Append("No subcategories defined.")
            End If
            Call NewLine(sbText)
            Call NewLine(sbText)
        End If
       
        sbText.Append(Bold("PRODUCT BREAKDOWN"))
        Call NewLine(sbText)
        Call NewLine(sbText)
        sSQL = "SELECT ProductCode, ISNULL(ProductDate, '') 'ProductDate', ProductDescription, ISNULL(ProductCategory, '') 'ProductCategory', ISNULL(SubCategory, '') 'SubCategory', ISNULL(SubCategory2, '') 'SubCategory2', ArchiveFlag FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey") & " ORDER BY ProductCode"
        dt = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                sbText.Append("Product code: ")
                sbText.Append(Bold(dr("ProductCode")))
                If dr("ProductCode").ToString.ToUpper = "Y" Then
                    sbText.Append(Bold("  (ARCHIVED)"))
                End If
                Call NewLine(sbText)
                If dr("ProductDate") = String.Empty Then
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold("(none)"))
                Else
                    sbText.Append("Product value/date: ")
                    sbText.Append(Bold(dr("ProductDate")))
                End If
                Call NewLine(sbText)
                sbText.Append("Product description: ")
                sbText.Append(Bold(dr("ProductDescription")))
                Call NewLine(sbText)
                sbText.Append("Category: ")
                sbText.Append(Bold(dr("ProductCategory")))
                Call NewLine(sbText)
                sbText.Append("Sub Category: ")
                sbText.Append(Bold(dr("SubCategory")))
                Call NewLine(sbText)
                If pnCategoryMode = CATEGORY_MODE_3_CATEGORIES Then
                    sbText.Append("Sub Sub Category: ")
                    sbText.Append(Bold(dr("SubCategory2")))
                    Call NewLine(sbText)
                End If
                Call NewLine(sbText)
            Next
        Else
            sbText.Append("No products defined.")
        End If
        Call NewLine(sbText)
        Call NewLine(sbText)

        sbText.Append("[end]")
        Call AddHTMLPostamble(sbText)
        Call ExportData(sbText.ToString, "CategoriesReport")
        Call WebMsgBox.Show("Category report downloaded.")
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Product Manager</title>
</head>
<body>
    <form id="frmProductManager" method="post" enctype="multipart/form-data" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr class="bar_productmanager">
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 50%; white-space: nowrap" align="right">
                <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="white"
                    Text="You have" />
                &nbsp;<asp:Label runat="server" ID="lblProductCount" ForeColor="#F9D938" Font-Names="Verdana"
                    Font-Size="XX-Small" Font-Bold="true"></asp:Label>
                &nbsp;<asp:Label ID="lblProductCountText" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    ForeColor="white" Text="live products and" />
                &nbsp;<asp:Label runat="server" ID="lblArchivedProductCount" ForeColor="#F9D938"
                    Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true"></asp:Label>
                &nbsp;<asp:Label ID="lblArchivedProductCountText" runat="server" Font-Names="Verdana"
                    Font-Size="XX-Small" ForeColor="white" Text="archived products" />&nbsp;&nbsp;
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlMainButtonRow" runat="server" Width="100%" Visible="true">
        <table width="100%" style="font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td align="left" valign="middle" style="white-space: nowrap">
                    <asp:Button ID="btn_ShowAllProducts" runat="server" OnClick="btn_ShowAllProducts_Click"
                        Text="show all products" ToolTip="get full product list" />
                    &nbsp;&nbsp;<asp:Button ID="btnShowCategories" runat="server" Text="show categories"
                        OnClick="btnShowCategories_Click" />
                    &nbsp;&nbsp;<asp:Label ID="Label19" runat="server" ForeColor="Gray" Font-Size="XX-Small"
                        Font-Names="Verdana">search:</asp:Label>
                    &nbsp;<asp:TextBox runat="server" Width="80px" Font-Size="XX-Small" ID="txtSearchCriteriaAllProducts"></asp:TextBox>
                    &nbsp;<asp:Button ID="btn_SearchAllProducts" OnClick="btn_SearchAllProducts_Click"
                        runat="server" Text="go" ToolTip="search across all products" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnAddProduct" OnClick="btn_AddProduct_click"
                        Text="add product" runat="server" />
                    &nbsp;&nbsp;&nbsp;<asp:Button ID="btnPendingAuthorisations" Visible="false" runat="server"
                        Text="show pending authorisations" OnClick="btnPendingAuthorisations_Click" />
                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:Button ID="btnProductGroups" runat="server"
                        Text="product groups" Visible="false" OnClick="btnProductGroups_Click" />
                    &nbsp;<asp:Button ID="btnProductCredits" runat="server" Text="product credits" Visible="false"
                        OnClick="btnProductCredits_Click" />
                </td>
                <td align="right" valign="middle">
                    <asp:LinkButton ID="lnkbtnCategoriesReport" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" onclick="lnkbtnCategoriesReport_Click">categories report</asp:LinkButton>
                </td>
                <td align="right" valign="middle" style="white-space: nowrap">
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlCategorySelection1" runat="server" Visible="True" Width="100%">
        <table id="tblCategorySelection" runat="server" width="100%" style="font-family: Verdana;
            font-size: small" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 2%">
                </td>
                <td valign="top" style="white-space: nowrap; width: 48%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label ID="Label111" runat="server" ForeColor="Navy" Font-Bold="True"
                        Font-Size="X-Small">Product Category</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" ID="Repeater1" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton5" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubCategories_click" CommandArgument='<%# Container.DataItem("Category")%>'
                                Text='<%# Container.DataItem("Category")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 48%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeading" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small">Sub Category</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater2">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton6" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowProductsByCategory_click" CommandArgument='<%# Container.DataItem("SubCategory")%>'
                                Text='<%# Container.DataItem("SubCategory")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td style="width: 2%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlCategorySelection2" runat="server" Visible="True" Width="100%">
        <table id="tblCategorySelection2" runat="server" width="100%" style="font-family: Verdana;
            font-size: small" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 2%">
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label ID="Label93" runat="server" ForeColor="Navy" Font-Bold="True"
                        Font-Size="X-Small">Product Category</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" ID="Repeater1a" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton1" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubCategories_click" CommandArgument='<%# Container.DataItem("Category")%>'
                                Text='<%# Container.DataItem("Category")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeadingA" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small">Sub Category</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater2a" OnItemCommand="repeater1_Item_click">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton3" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowSubSubCategories_click" CommandArgument='<%# Container.DataItem("SubCategory")%>'
                                Text='<%# Container.DataItem("SubCategory")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td valign="top" style="white-space: nowrap; width: 32%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label runat="server" ID="lblSubCategoryHeadingB" ForeColor="Navy"
                        Font-Bold="True" Font-Size="X-Small" visible="false">Sub Category 2</asp:Label>
                    <br />
                    <br />
                    <asp:Repeater runat="server" Visible="False" ID="Repeater3a">
                        <ItemTemplate>
                            &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="LinkButton4" runat="server" Style="text-decoration: none"
                                OnCommand="lnkbtnShowProductsByCategory_click" CommandArgument='<%# Container.DataItem("SubCategory2")%>'
                                Text='<%# Container.DataItem("SubCategory2")%>' ForeColor="Navy" Font-Size="X-Small"
                                EnableTheming="false" />
                            <br />
                        </ItemTemplate>
                    </asp:Repeater>
                    <br />
                </td>
                <td style="width: 2%">
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlProductList" runat="server" Visible="False" Width="100%">
        <asp:DataGrid ID="dg_ProductList" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" PageSize="6" OnPageIndexChanged="dg_ProductList_Page_Change"
            AllowPaging="True" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" OnItemCommand="dg_ProductList_item_click">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <PagerStyle NextPageText="Next Page  " Font-Size="X-Small" Font-Names="Verdana" Font-Bold="True"
                PrevPageText="Previous Page" HorizontalAlign="Center" ForeColor="Blue" Position="Top"
                BackColor="Silver" Wrap="False" Mode="NumericPages"></PagerStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="LogisticProductKey">
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <table id="tabProductList" runat="server" style="font: Verdana; font-size: XX-Small;
                            color: Gray; width: 100%">
                            <tr>
                                <td rowspan="4" valign="top" style="width: 7%">
                                    <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'></asp:HyperLink>
                                </td>
                                <td style="width: 12%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label5" runat="server" ForeColor="Gray">Product Code:</asp:Label>
                                </td>
                                <td style="width: 15%" valign="top" wrap="False">
                                    <asp:Label ID="Label4" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductCode") %></asp:Label>
                                </td>
                                <td style="width: 12%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label6" runat="server" ForeColor="Gray" Text="Version/Date:" />
                                </td>
                                <td style="width: 15%" valign="top">
                                    <asp:Label ID="Label7" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductDate") %></asp:Label>
                                </td>
                                <td style="width: 27%">
                                </td>
                                <td style="width: 12%; white-space: nowrap" valign="top" align="right">
                                    <asp:Label ID="Label8" runat="server" ForeColor="Gray">Quantity:</asp:Label>
                                    &nbsp;<asp:Label ID="Label9" runat="server" ForeColor="Navy"><%# Format(DataBinder.Eval(Container.DataItem,"Quantity"),"#,##0") %></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <asp:Label ID="Label10" runat="server" ForeColor="Gray">Category:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label11" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ProductCategory") %></asp:Label>
                                </td>
                                <td valign="top" wrap="False">
                                    <asp:Label ID="Label12" runat="server" ForeColor="Gray">Sub Category:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label13" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"SubCategory") %></asp:Label>
                                </td>
                                <td valign="top" align="right">
                                    <asp:Label ID="Label16" runat="server" ForeColor="Gray" Text="Archive Flag:" />
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label17" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ArchiveFlag") %></asp:Label>
                                </td>
                            </tr>
                            <tr id="trJupiterPODType" runat="server" visible="<%# IsJupiterPODProduct(Container.DataItem) %>">
                                <td valign="top" style="white-space: nowrap">
                                    <asp:Label ID="lblLegendPrintType2" runat="server" ForeColor="#FF3300">Print Type:</asp:Label>
                                </td>
                                <td valign="top" colspan="4" rowspan="2">
                                    <asp:Label ID="lblPrintType2" runat="server" ForeColor="#FF3300" Font-Bold="True"><%# GetJupiterPODProductType(Container.DataItem) %></asp:Label>
                                </td>
                                <td valign="bottom" align="right" rowspan="2">
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="8" valign="top">
                                </td>
                            </tr>
                            <tr>
                                <td valign="top" style="white-space: nowrap">
                                    <asp:Label ID="Label14" runat="server" ForeColor="Gray">Description:</asp:Label>
                                </td>
                                <td valign="top" colspan="4" rowspan="2">
                                    <asp:Label ID="Label15" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ProductDescription") %></asp:Label>
                                </td>
                                <td valign="bottom" align="right" rowspan="2">
                                    <asp:Button ID="EditProduct" runat="server" CommandName="Edit" Text="edit this product"
                                        ToolTip="edit this product" />
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="8" valign="top">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid><asp:Label ID="Label72" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            ForeColor="Gray">Items per page:</asp:Label>&nbsp;<asp:DropDownList ID="ddlItemsPerPage"
                runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small"
                OnSelectedIndexChanged="ddlItemsPerPage_SelectedIndexChanged">
                <asp:ListItem Selected="True">6</asp:ListItem>
                <asp:ListItem>20</asp:ListItem>
                <asp:ListItem>50</asp:ListItem>
            </asp:DropDownList>
        <asp:Label ID="lblProductMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="Red" Font-Bold="True"></asp:Label></asp:Panel>
    <asp:Panel ID="pnlEditProduct" runat="server" Visible="False" Width="100%">
        <table id="table1xx" width="100%" style="font-family: Verdana; font-size: x-small">
            <tr valign="middle">
                <td style="white-space: nowrap; width: 40%; height: 26px;">
                    <asp:Label ID="lblLegendProductDetail" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                        Font-Bold="True" ForeColor="Gray">Product Detail: </asp:Label>&nbsp;<asp:Label runat="server"
                            ID="lblProductQuantity" Font-Size="X-Small" Font-Names="Verdana" ForeColor="Red">
                        </asp:Label><asp:Label ID="lblLegendItemsInStock" runat="server" Font-Size="X-Small"
                            Font-Names="Verdana" ForeColor="Gray"> items in stock.</asp:Label>
                </td>
                <td align="right" style="white-space: nowrap; width: 60%; height: 26px;">
                    <asp:LinkButton ID="lnkbtnShowHelp" runat="server" OnClick="lnkbtnShowHelp_Click"
                        CausesValidation="False">hide help</asp:LinkButton>&nbsp; &nbsp;<asp:Button ID="btnAssociatedProducts"
                            runat="server" OnClick="btnAssociatedProducts_Click" Text="associated products..." />
                    &nbsp;<asp:Button ID="btnSetUserProfiles" runat="server" Text="set max order levels..."
                        OnClick="btnSetUserProfiles_click" ToolTip="set max order levels" />&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button
                            ID="btn_DeleteProduct" runat="server" OnClick="btn_DeleteProduct_click" OnClientClick="return confirm(&quot;Are you sure you want to delete this product?&quot;);"
                            Text="delete product" ToolTip="delete this product" /><a id="aHelpDeleteProduct"
                                runat="server" onmouseover="return escape('Click this button to remove the product completely. As a precaution against accidental deletion you must click the OK button when asked \'Are you sure?\'. To be deleted the product must have a stock level of 0 (zero). You can restore a deleted product using the UnDelete facility (not currently enabled).')"
                                style="color: gray; cursor: help; font-size: xx-small" visible="false">&nbsp;?&nbsp;</a>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btn_GoToProductListPanel" runat="server"
                        OnClick="btn_GoToProductListPanel_click" CausesValidation="false" Text="return to list"
                        ToolTip="go back to product list" />
                </td>
            </tr>
        </table>
        <br />
        <table style="width: 100%">
            <tr>
                <td style="width: 12%; white-space: nowrap">
                </td>
                <td style="width: 25%; white-space: nowrap">
                </td>
                <td style="width: 12%; white-space: nowrap">
                </td>
                <td style="width: 24%; white-space: nowrap">
                </td>
                <td style="width: 15%; white-space: nowrap">
                </td>
                <td style="width: 12%; white-space: nowrap">
                </td>
            </tr>
            <tr>
                <td colspan="6" align="right">
                    <asp:Label runat="server" ForeColor="#00C000" ID="lblEditDateError" Font-Size="X-Small"
                        Font-Names="Verdana"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdProductCode" runat="server" ControlToValidate="txtProductCode"
                        Font-Size="XX-Small" Text="#" />
                    &nbsp;
                    <asp:Label ID="lblLegendProdCode" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Prod Code:" />
                </td>
                <td style="white-space: nowrap">
                    <asp:TextBox ID="txtProductCode" MaxLength="25" runat="server" ForeColor="Navy" Width="150"
                        TabIndex="1" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpProductCode" visible="false" onmouseover="return escape('<b>Product Code</b> (maximum length 25 chars) when combined with <b>Product Date</b> (sometimes called <b>Version Date</b>) uniquely identifies this product.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                    <asp:LinkButton ID="lnkbtnNewProductCode" Visible="false" runat="server" CausesValidation="False"
                        Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnNewProductCode_Click">new code</asp:LinkButton><a
                            runat="server" id="aHelpNewCode" visible="false" onmouseover="return escape('Click <b>new code</b> to assign a new, unique product code to this product.<br /><br />If you require a new, unique product code but you do not want to create the product at this time (eg because the Version Date is not yet available) use the <b>reserve W-code</b> facility.<br /><br />When you subsequently create the product, enter the product code you reserved. That code will then be removed from your reservation list.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdProductDate" runat="server" ControlToValidate="txtProductDate"
                        Enabled="false" Font-Size="XX-Small" Text="#" />
                    &nbsp;
                    <asp:Label ID="lblLegendProductDate" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Text="Product Date:" />
                </td>
                <td>
                    <asp:TextBox ID="txtProductDate" MaxLength="10" runat="server" ForeColor="Navy" Width="100"
                        TabIndex="2" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpProductDate" visible="false" onmouseover="return escape('<b>Product Date</b> (sometimes called <b>Version Date</b>) when combined with <b>Product Code</b> uniquely identifies this product. Use this field to identify a specific version or variant of a product, the versions of which share the same <b>Product Code</b>. Maximum length 10 chars.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:RegularExpressionValidator ID="revMinStockLevel" runat="server" ControlToValidate="txtMinStockLevel"
                        Enabled="False" Font-Size="XX-Small" ValidationExpression="[123456789]\d*">#</asp:RegularExpressionValidator><asp:RequiredFieldValidator
                            ID="rfdMinStockLevel" runat="server" ControlToValidate="txtMinStockLevel" Font-Size="XX-Small"
                            Text="#" Enabled="False" />
                    &nbsp;
                    <asp:Label ID="lblLegendMinStockLevel" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Min Stock Level:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtMinStockLevel" runat="server" ForeColor="Navy" Width="50" TabIndex="3"
                        Font-Size="XX-Small" Font-Names="Verdana" MaxLength="6"></asp:TextBox><a runat="server"
                            id="aHelpMinStockLevel" visible="false" onmouseover="return escape('The system sends an email alert when the available stock quantity falls to (or below) this level. In some installations this field is mandatory. If this field is not mandatory, you can set the value to 0 to disable <b>Low Stock</b> email alerts.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdDescription" runat="server" ControlToValidate="txtDescription"
                        Font-Size="XX-Small" Text="#" Enabled="False" />
                    &nbsp;
                    <asp:Label ID="lblLegendDescription" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Description:</asp:Label>
                </td>
                <td colspan="3">
                    <asp:TextBox ID="txtDescription" MaxLength="300" runat="server" ForeColor="Navy"
                        Width="470px" TabIndex="4" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a
                            runat="server" id="aHelpDescription" visible="false" onmouseover="return escape('Description of the product. Maximum length 300 characters.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendItemsPerBox" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Items Per Box:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtItemsPerBox" runat="server" ForeColor="Navy" Width="50" TabIndex="5"
                        Font-Size="XX-Small" Font-Names="Verdana" MaxLength="6"></asp:TextBox><a runat="server"
                            id="aHelpItemsPerBox" visible="false" onmouseover="return escape('The number of individual pieces of this product in a box (or other container). Typically used to give guidance on preferred order quantities. Unless explicitly stated, the available stock quantity refers to the number of individual pieces of a product, not the number of boxes.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr id="trPrintProperties" runat="server" visible="false">
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendOnDemand2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">On Demand:</asp:Label>
                </td>
                <td colspan="5">
                    <asp:CheckBox ID="cbOnDemand" runat="server" Font-Italic="False" Font-Names="Verdana"
                        Font-Size="XX-Small" AutoPostBack="True" OnCheckedChanged="cbOnDemand_CheckedChanged" />
                    &nbsp;<asp:Label ID="lblPODPrintType" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small">(undefined)</asp:Label>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfdPODPageCount" runat="server" ControlToValidate="ddlPODPageCount" Enabled="False" Font-Bold="True" Font-Size="XX-Small" ForeColor="Red" InitialValue="0" Text="#" Visible="False" />
                    &nbsp;<asp:Label ID="lblLegendPODPageCount" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Visible="False">Pages:</asp:Label>
                    <%--<asp:DropDownList ID="ddlPrintType" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlPrintType_SelectedIndexChanged" Visible="False"/>--%>
                    &nbsp;<asp:DropDownList ID="ddlPODPageCount" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" onselectedindexchanged="ddlPODPageCount_SelectedIndexChanged" Visible="False">
                        <asp:ListItem Selected="True" Value="0">- please select -</asp:ListItem>
                        <asp:ListItem>1</asp:ListItem>
                        <asp:ListItem>2</asp:ListItem>
                        <asp:ListItem>4</asp:ListItem>
                        <asp:ListItem>8</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>32</asp:ListItem>
                        <asp:ListItem>36</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="lblLegendPODSize" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Size:</asp:Label>
                    <asp:RadioButton ID="rbPODSizeA4" runat="server" Checked="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="PODStockSize" oncheckedchanged="rbPOD_CheckedChanged" Text="A4" AutoPostBack="True" />
                    <asp:RadioButton ID="rbPODSizeA5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="PODStockSize" oncheckedchanged="rbPOD_CheckedChanged" Text="A5" AutoPostBack="True" />
                    &nbsp; &nbsp;
                    <asp:Label ID="lblLegendPODStock" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Stock:</asp:Label>
                    <asp:RadioButton ID="rbPOD120gsm" runat="server" AutoPostBack="True" Checked="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="PODStockWeight" oncheckedchanged="rbPOD_CheckedChanged" Text="120 gsm" />
                    <asp:RadioButton ID="rbPOD150gsm" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="PODStockWeight" oncheckedchanged="rbPOD_CheckedChanged" Text="150 gsm" />
                    <asp:RadioButton ID="rbPOD200gsm" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="PODStockWeight" oncheckedchanged="rbPOD_CheckedChanged" Text="200 gsm" />
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendCategory" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Category:</asp:Label>
                </td>
                <td>
                    <asp:HiddenField ID="hidCategory" runat="server" />
                    <asp:TextBox ID="txtCategory" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="6" Visible="False" Width="150" /><asp:DropDownList
                            ID="ddlCategory" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceCategoryList"
                            DataTextField="Category" DataValueField="Category" EnableViewState="false" Font-Names="Verdana"
                            Font-Size="XX-Small" OnDataBound="ddlCategory_DataBound" OnSelectedIndexChanged="ddlCategory_SelectedIndexChanged"
                            Visible="False" Width="150">
                        </asp:DropDownList>
                    <a id="aHelpCategory" runat="server" onmouseover="return escape('The top level category for this product. Add a new category by clicking &lt;b&gt;- new category -&lt;/b&gt; then entering the new category name. If you change the top level category of a product, you must then set the sub category.')"
                        style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendProductUsersMessage" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Sub Category:</asp:Label>
                </td>
                <td>
                    <asp:HiddenField ID="hidSubCategory" runat="server" />
                    <asp:TextBox ID="txtSubCategory" MaxLength="50" runat="server" ForeColor="Navy" TabIndex="7"
                        Width="150" Font-Size="XX-Small" Font-Names="Verdana" Visible="False"></asp:TextBox><asp:DropDownList
                            ID="ddlSubCategory" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceSubCategoryList"
                            DataTextField="SubCategory" DataValueField="SubCategory" EnableViewState="false"
                            Font-Names="Verdana" Font-Size="XX-Small" OnDataBound="ddlSubCategory_DataBound"
                            OnSelectedIndexChanged="ddlSubCategory_SelectedIndexChanged" Visible="False"
                            Width="150">
                        </asp:DropDownList>
                    <a runat="server" id="aHelpSubCategory" visible="false" onmouseover="return escape('The 2nd level (sub) category for this product.  Add a new sub category by clicking &lt;b&gt;- new subcategory -&lt;/b&gt; then entering the new sub category name'). If you are using a further sub category level, you must then set the further (final) sub category."
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendArchiveFlag" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Archive Flag:</asp:Label>
                </td>
                <td>
                    <asp:CheckBox ID="chkArchivedFlag" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        TabIndex="14" /><a id="aHelpArchived" runat="server" onmouseover="return escape('Controls whether this product is shown on the Orders page')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr id="trSubCategory2" runat="server" visible="false">
                <td align="right">
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendSubCategory2" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Sub Category 2:</asp:Label>
                </td>
                <td>
                    <asp:HiddenField ID="hidSubSubCategory" runat="server" />
                    <asp:TextBox ID="tbSubSubCategory" MaxLength="50" runat="server" ForeColor="Navy"
                        TabIndex="7" Width="150" Font-Size="XX-Small" Font-Names="Verdana" Visible="False"></asp:TextBox><asp:DropDownList
                            ID="ddlSubSubCategory" runat="server" AutoPostBack="True" DataSourceID="SqlDataSourceSubSubCategoryList"
                            DataTextField="SubCategory2" DataValueField="SubCategory2" EnableViewState="false"
                            Font-Names="Verdana" Font-Size="XX-Small" OnDataBound="ddlSubSubCategory_DataBound"
                            OnSelectedIndexChanged="ddlSubSubCategory_SelectedIndexChanged" Visible="True"
                            Width="150">
                        </asp:DropDownList>
                    <a runat="server" id="aHelpSubSubCategory" visible="false" onmouseover="return escape('The 3rd level (sub) category for this product.  Add a new sub category by clicking &lt;b&gt;- new subcategory -&lt;/b&gt; then entering the new sub category name').')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendLanguage" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Language:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtLanguage" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="9" Width="150"></asp:TextBox><a id="aHelpLanguage"
                            runat="server" onmouseover="return escape('The language of this product')" style="color: gray;
                            cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:RequiredFieldValidator ID="rfdCostCentre" runat="server" ControlToValidate="txtDepartment"
                        Enabled="False" Font-Size="XX-Small" Text="#" />&nbsp;
                    <asp:Label ID="lblLegendCostCentre" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Cost Centre:</asp:Label>
                </td>
                <td style="height: 22px">
                    <asp:TextBox ID="txtDepartment" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="10" Width="150"></asp:TextBox><a id="aHelpDepartment"
                            runat="server" onmouseover="return escape('The cost centre associated with this product')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendUnitWeight" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Unit Weight (gm):</asp:Label>
                </td>
                <td style="height: 22px">
                    <asp:TextBox ID="txtUnitWeight" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="6" TabIndex="11" Width="50"></asp:TextBox><a id="aHelpUnitWeight"
                            runat="server" onmouseover="return escape('The weight in grams of a single unit of this product')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap;">
                    <asp:Label ID="lblLegendMisc1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Misc 1:"></asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtMisc1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="12" Width="150"></asp:TextBox><a id="aHelpMisc1"
                            runat="server" onmouseover="return escape('This field is available to store additional data of your choice. Data in this field will be checked when using the &lt;b&gt;search&lt;/b&gt; facility. Maximum length 50 characters.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>&nbsp;
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:RegularExpressionValidator ID="revFEXCOCriticalProduct" runat="server" ControlToValidate="txtMisc2"
                        Enabled="False" Font-Size="XX-Small" ValidationExpression="[YyNn]">#</asp:RegularExpressionValidator>&nbsp;
                    <asp:Label ID="lblLegendMisc2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Misc 2:" />
                </td>
                <td>
                    <asp:TextBox ID="txtMisc2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="13" Width="150"></asp:TextBox><a id="aHelpMisc2"
                            runat="server" onmouseover="return escape('This field is available to store additional data of your choice. Data in this field will be checked when using the &lt;b&gt;search&lt;/b&gt; facility. Maximum length 50 characters.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendUnitValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Unit Value (£):</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtUnitValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="6" TabIndex="8" Width="50"></asp:TextBox><a id="aHelpUnitValue"
                            runat="server" onmouseover="return escape('The value, or cost price, of a single item or unit of this product')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr id="trOwner" runat="server" visible="false">
                <td align="right" style="white-space: nowrap">
                    <asp:RequiredFieldValidator ID="rfdProductGroup" runat="server" ControlToValidate="ddlAssignedProductGroup"
                        Enabled="False" Font-Size="XX-Small" InitialValue="0" Text="#" />&nbsp;
                    <asp:Label ID="lblLegendProductGroup" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Prod Group:</asp:Label>
                </td>
                <td style="white-space: nowrap;" colspan="2">
                    <asp:DropDownList ID="ddlAssignedProductGroup" runat="server" AutoPostBack="True"
                        Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlAssignedProductGroup_SelectedIndexChanged"
                        Visible="True" Width="160" />
                    &nbsp;<asp:Label ID="lblAssignedProductOwners" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small"></asp:Label><a runat="server" id="aHelpAssignedProductGroup"
                            visible="false" onmouseover="return escape('The Product Group for this product, used to assign Stock Owners to Products.&lt;br /&gt;&lt;br /&gt; The name of the Primary Stock Owner and/or Deputy Stock Owner will appear next to the Product Group name, if these roles have been assigned.&lt;br /&gt;&lt;br /&gt; Product Groups are created, and Stock Owners assigned, by Super Users.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="right" style="white-space: nowrap;">
                    <asp:Label ID="lblLegendSellingPrice" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Visible="false">Selling Price (£):</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbSellingPrice" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="6" TabIndex="8" Visible="false" Width="50" /><a id="aHelpSellingPrice"
                            runat="server" onmouseover="return escape('The selling price of a single item or unit of this product')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendExpiryDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Expiry Date:</asp:Label>
                </td>
                <td style="height: 37px">
                    <asp:TextBox ID="tbExpiryDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="90" /><a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbExpiryDate','cal','width=300,height=305,left=270,top=180')"><img
                            id="Img1" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0" ie:visible="true"
                            visible="false" /></a><a runat="server" id="aHelpExpiryDate" visible="false" onmouseover="return escape('When the Expiry Date is reached or passed, the system periodically sends an email alert. To activate this function click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date , or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2006 ')"
                                style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendRenewalDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Renewal Date:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbReplenishmentDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Width="90" /><a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=tbReplenishmentDate','cal','width=300,height=305,left=270,top=180')"><img
                            id="SmallCalendar" alt="" src="~/images/SmallCalendar.gif" runat="server" border="0"
                            ie:visible="true" visible="false" /></a><a id="aHelpReplenishmentDate" runat="server"
                                onmouseover="return escape('When the Renewal (sometimes called Review or Replenishment) Date is reached or passed, the system periodically sends an email alert. To activate this function click the calendar icon (available in Internet Explorer only) and follow the instructions to select a date, or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2006 ')"
                                style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>&nbsp;
                </td>
                <td align="right" style="white-space: nowrap; height: 37px;">
                    <asp:Label ID="lblLegendSerialNumbers" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Serial Numbers:</asp:Label>
                </td>
                <td id="CHECK_THIS_FIELD" style="height: 37px">
                    <asp:CheckBox ID="chkProspectusNumbers" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        TabIndex="21" /><a id="aHelpProspectusNumbers" runat="server" onmouseover="return escape('Controls whether serial numbers are used for this product. Serial numbers are typically used to track products for auditing.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>&nbsp;
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td align="right" style="white-space: nowrap">
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendInactivityAlert" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Product Inactivity Alert:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbInactivityAlertDays" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="3" TabIndex="3" Width="50" /><asp:LinkButton ID="lnkbtnConfigureProductInactivityAlerts"
                            runat="server" CausesValidation="False" Font-Names="Verdana" Font-Size="XX-Small"
                            OnClick="lnkbtnConfigureProductInactivityAlerts_Click">config</asp:LinkButton><a
                                id="aHelpInactivityAlert" runat="server" onmouseover="return escape('Sets the period, in days, after which an email alert will be sent if no orders have been received for this product. A value of zero (0) disables this feature.&lt;br/&gt;&lt;br/&gt;Click the &lt;b&gt;config&lt;/b&gt; button to configure Inactivity Alerts for all products.')"
                                style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>&nbsp;
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendCustomLetter" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Custom letter:</asp:Label>
                </td>
                <td>
                    <asp:CheckBox ID="cbCustomLetter" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnCheckedChanged="cbCustomLetter_CheckedChanged" TabIndex="23" /><a
                            id="aHelpCustomLetter" runat="server" onmouseover="return escape('Indicates if this product is a custom letter, the text of which can be specified when the product is ordered')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>&nbsp;<asp:LinkButton
                                ID="lnkbtnConfigureCustomLetter" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                OnClick="lnkbtnConfigureCustomLetter_Click">configure custom letter</asp:LinkButton>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendProductCredits" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Product credits:</asp:Label>
                </td>
                <td>
                    <asp:CheckBox ID="cbProductCredits" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnCheckedChanged="cbProductCredits_CheckedChanged" />&nbsp;<a
                            id="aHelpProductCredits" runat="server" onmouseover="return escape('Indicates if allocation of this product is controlled by product credits')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                    <asp:LinkButton ID="lnkbtnConfigureProductCredits" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" OnClick="lnkbtnConfigureProductCredits_Click">configure product credits</asp:LinkButton>
                </td>
                <td align="right" style="white-space: nowrap">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap;">
                    &nbsp;
                </td>
                <td>
                </td>
                <td align="right" style="white-space: nowrap;">
                    &nbsp;
                </td>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr id="trAdRotatorControls" runat="server" visible="true">
                <td align="right" style="white-space: nowrap;">
                    <asp:Label ID="lblLegendAdRotatorText" runat="server" Font-Names="Verdana" Font-Size="XX-Small">AdRotator Txt:</asp:Label>
                </td>
                <td colspan="3" valign="top">
                    <asp:TextBox ID="txtAdRotatorText" MaxLength="120" Width="420px" runat="server" ForeColor="Navy"
                        TabIndex="22" Font-Size="XX-Small" Font-Names="Verdana"></asp:TextBox><a runat="server"
                            id="aHelpAdRotatorText" visible="false" onmouseover="return escape('The text to appear in the \'advertisement\' above the tabs in the browser window')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap;">
                    <asp:Label ID="lblLegendAdRotator" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Ad Rotator:</asp:Label>
                </td>
                <td>
                    <asp:CheckBox runat="server" ID="chkAdRotator" Font-Names="Verdana" Font-Size="XX-Small"
                        TabIndex="23" /><a runat="server" id="aHelpAdRotator" visible="false" onmouseover="return escape('Controls whether \'advertisements\' are displayed above the tabs in the browser window')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr>
                <td align="right" style="white-space: nowrap; height: 22px;" valign="top">
                    <asp:Label ID="lblLegendComments" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Comments:</asp:Label>
                </td>
                <td colspan="3" rowspan="2">
                    <asp:TextBox ID="txtNotes" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="1000" Rows="3" TabIndex="24" TextMode="MultiLine"
                        Width="420"></asp:TextBox><a id="aHelpNotes" runat="server" onmouseover="return escape('Additional notes on this product. Depending on your installation options, these notes may appear on the Order page to convey additional information about the product to orderers.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td align="right" style="white-space: nowrap; height: 22px;">
                    <asp:Label ID="lblLegendViewOnWebForm" runat="server" Font-Names="Verdana" Font-Size="XX-Small">View on Web Form:</asp:Label>
                </td>
                <td style="height: 22px">
                    <asp:CheckBox ID="chkViewOnWebForm" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        TabIndex="25" /><a id="aHelpViewOnWebForm" runat="server" onmouseover="return escape('Controls whether this product is displayed on additional web forms. Web Forms are an installation option.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="lblLegendRequiresAuth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Requires Auth:</asp:Label>
                </td>
                <td>
                    <asp:CheckBox ID="cbRequiresAuth" runat="server" AutoPostBack="true" Font-Names="Verdana"
                        Font-Size="XX-Small" OnCheckedChanged="cbRequiresAuth_CheckedChanged" TabIndex="25" /><asp:LinkButton
                            ID="lnkbtnPreAuthorise" runat="server" CausesValidation="False" Font-Names="Verdana"
                            Font-Size="XX-Small" OnClick="lnkbtnPreAuthorise_Click">pre-authorise</asp:LinkButton><span
                                id="spanProductAuthorisationAd" runat="server" visible="true"><a id="tooltipRequiresAuth"
                                    runat="server" onmouseover="return escape('Product Authorisation gives you fine control over product ordering by users.  This feature is currently disabled. To enable Product Authorisation contact your Account Handler.')"
                                    style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a></span>
                </td>
            </tr>
            <tr valign="top">
                <td align="right" rowspan="2">
                    <asp:HyperLink ID="hlnk_DetailThumb" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Target="_blank" ToolTip="click here to see larger image" />
                </td>
                <td>
                    <input id="fuBrowseImageFile" style="width: 240px; font-family: Verdana; font-size: xx-small"
                        type="file" runat="server" /><asp:Label ID="lblImageUploadUnavailable" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text="(image upload unavailable until product created)"></asp:Label>
                </td>
                <td align="right" rowspan="2">
                    <asp:HyperLink ID="hlnk_PDF" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Target="_blank" ToolTip="click here to view PDF file" />
                </td>
                <td>
                    <input id="fuBrowsePDFFile" style="width: 240px; font-family: Verdana; font-size: xx-small"
                        type="file" runat="server" enableviewstate="False" /><asp:Label ID="lblPDFUploadUnavailable" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text="(PDF upload unavailable until product created)"/>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendViewOnWebFormDE" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="View on Web Form DE:" Visible="False" /><asp:Label ID="lblLegendCalendarManaged"
                            runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Calendar Managed:"
                            Visible="false" />
                </td>
                <td>
                    <asp:CheckBox ID="cbViewOnWebFormDE" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        TabIndex="25" Visible="false" /><asp:CheckBox ID="cbCalendarManaged" runat="server"
                            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbCalendarManaged_CheckedChanged"
                            TabIndex="25" Visible="true" /><a id="aHelpCalendarManaged" runat="server" onmouseover="return escape('Set a product as Calendar Managed to add support for delivery and collection. This is typically used for products such as exhibition equipment where the same items is used at successive locations. Contact your Account Handler for further details of this facility.')"
                                style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
            </tr>
            <tr valign="top">
                <td>
                    <asp:Button ID="btnUploadImage" runat="server" CausesValidation="False" OnClick="btnUploadImage_click"
                        Text="upload jpg" ToolTip="upload the selected jpg file to the server" />&nbsp;
                    <asp:ImageButton ID="imgbtnDeleteImage" runat="server" ImageUrl="~/images/delete.gif"
                        OnClick="imgbtnDeleteImage_Click" OnClientClick="return confirm(&quot;Are you sure you want to delete this image?&quot;);"
                        ToolTip="delete this image" /><a id="aHelpUploadImage" runat="server" onmouseover="return escape('Allows you to upload a picture of this product. The picture must be in standard JPG format.  Pictures are automatically resized on upload if necessary.')"
                            style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td>
                    <asp:Button ID="btnUploadPDF" runat="server" CausesValidation="False" OnClick="btnUploadPDF_click"
                        Text="upload pdf" ToolTip="upload the selected pdf file to the server"  Visible='<%# SetUploadPDFVisibility() %>' /><asp:ImageButton
                            ID="imgbtnDeletePDF" runat="server" ImageUrl="~/images/delete.gif" OnClick="imgbtnDeletePDF_Click"
                            OnClientClick="return confirm(&quot;Are you sure you want to delete this PDF?&quot;);"
                            ToolTip="delete this PDF" /><a id="aHelpUploadPDF" runat="server" onmouseover="return escape('Allows you to upload an Adobe PDF file which can be downloaded by orderers eg to provide further information about a product')"
                                style="color: gray; cursor: help" visible="false">&nbsp;?&nbsp;</a>
                </td>
                <td>
                </td>
                <td align="right">
                    <asp:Button ID="btn_SaveProductChanges" runat="server" OnClick="btn_SaveProductChanges_click"
                        Text="save changes" ToolTip="save changes to product record" />&nbsp;&nbsp;
                </td>
            </tr>
        </table>
        <asp:RegularExpressionValidator ID="revExpiryDate" runat="server" ErrorMessage="Invalid format for expiry date - use dd-mmm-yyyy"
            ControlToValidate="tbExpiryDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
            Font-Names="Verdana" Font-Size="X-Small" EnableClientScript="false"></asp:RegularExpressionValidator><br />
        <asp:RangeValidator ID="rvExpiryDate" runat="server" ErrorMessage="Expiry year before 2000, after 2020, or not a valid date!"
            EnableClientScript="false" Enabled="false" ControlToValidate="tbExpiryDate" MaximumValue="2019/1/1"
            MinimumValue="2000/1/1" CultureInvariantValues="True" Font-Names="Verdana" Font-Size="X-Small">
        </asp:RangeValidator><br />
        <asp:RegularExpressionValidator ID="revReplenishmentDate" runat="server" ErrorMessage="Invalid format for renewal date - use dd-mmm-yyyy"
            ControlToValidate="tbReplenishmentDate" ValidationExpression="^\d\d-(jan|Jan|feb|Feb|mar|Mar|apr|Apr|may|May|jun|Jun|jul|Jul|aug|Aug|sep|Sep|oct|Oct|nov|Nov|dec|Dec)-\d\d\d\d"
            Font-Names="Verdana" Font-Size="X-Small" EnableClientScript="false"></asp:RegularExpressionValidator><br />
        <asp:RangeValidator ID="rvReplenishmentDate" runat="server" ErrorMessage="Renewal year before 2000, after 2020, or not a valid date!"
            EnableClientScript="false" Enabled="false" ControlToValidate="tbReplenishmentDate"
            MaximumValue="2019/1/1" MinimumValue="2000/1/1" CultureInvariantValues="True"
            Font-Names="Verdana" Font-Size="X-Small">
        </asp:RangeValidator><asp:Label ID="lblDateError" runat="server" Font-Names="Verdana"
            Font-Size="X-Small" Text="Please check Expiry Date and Replenishment Date. If present, these must be valid dates and the year must be between 2000 and 2020"
            ForeColor="Red" Visible="false" />
    </asp:Panel>
    <asp:SqlDataSource ID="SqlDataSourceCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSourceSubCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetSubCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
            <asp:ControlParameter ControlID="ddlCategory" Name="Category" PropertyName="SelectedValue"
                Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="SqlDataSourceSubSubCategoryList" runat="server" ConnectionString="<%$ ConnectionStrings:AIMSRootConnectionString %>"
        SelectCommand="spASPNET_Product_GetSubSubCategoriesIncludeArchivedProds" SelectCommandType="StoredProcedure">
        <SelectParameters>
            <asp:SessionParameter Name="CustomerKey" SessionField="CustomerKey" Type="Int32" />
            <asp:ControlParameter ControlID="ddlSubCategory" Name="SubCategory" PropertyName="SelectedValue"
                Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:Panel ID="pnlAuthoriseProduct" runat="server" Visible="False" Width="100%" Font-Names="Verdana"
        Font-Size="X-Small">
        <table width="95%">
            <tr>
                <td style="width: 80%">
                    <asp:Label ID="Label24aa" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="User Authorisation"></asp:Label>
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnAuthGoBack2" runat="server" Text="go back" CausesValidation="false"
                        OnClick="btn_GoToProductListPanel_click" />
                </td>
            </tr>
        </table>
        <br />
        The following products await authorisation.<br />
        <br />
        Click the Authorise check box for each request you want to authorise. Enter the
        quantity you want to authorise. You can optionally limit the period of validity
        for the authorisation by entering the duration in days that authorisation should
        last. If you don't wish to limit it, leave the duration as 'unlimited' or blank.<br />
        <br />
        Set the 'Leave items pending...' or 'Decline authorisation...' option depending
        on your preference. Finally, click the Authorise button.<br />
        <br />
        <asp:GridView ID="gvAuthoriseProduct" runat="server" AutoGenerateColumns="False"
            AllowPaging="True" Font-Names="Verdana" Font-Size="XX-Small" ShowFooter="True"
            Width="95%" GridLines="None">
            <Columns>
                <asp:BoundField DataField="id" HeaderText="id" InsertVisible="False" ReadOnly="True"
                    SortExpression="id" Visible="False" />
                <asp:BoundField DataField="LogisticProductKey" HeaderText="LogisticProductKey" InsertVisible="False"
                    ReadOnly="True" SortExpression="LogisticProductKey" Visible="False" />
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" SortExpression="ProductCode" />
                <asp:BoundField DataField="ProductDate" HeaderText="Product Date" SortExpression="ProductDate" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" SortExpression="ProductDescription" />
                <asp:BoundField DataField="QtyInStock" HeaderText="Qty In Stock" ReadOnly="True"
                    SortExpression="QtyInStock" />
                <asp:TemplateField HeaderText="Qty Requested" SortExpression="RequestedQty">
                    <HeaderTemplate>
                        &nbsp;Qty Auth'd<br />
                        (Qty Req'd)</HeaderTemplate>
                    <ItemTemplate>
                        &nbsp;<asp:HiddenField ID="hidAuthoriseId" Value='<%# Bind("id") %>' runat="server" />
                        &nbsp;&nbsp;<asp:TextBox ID="tbQuantityAuthorised" BackColor="LightYellow" runat="server"
                            Font-Size="XX-Small" Font-Names="Verdana" Text='<%# Bind("RequestedQty") %>'
                            Width="36px" />&nbsp;(<asp:Label ID="lblQtyRequested" runat="server" Text='<%# Bind("RequestedQty") %>'></asp:Label>)
                        <asp:RequiredFieldValidator ID="rfdQuantityAuthorised" runat="server" ControlToValidate="tbQuantityAuthorised"
                            EnableClientScript="False" ErrorMessage="*"></asp:RequiredFieldValidator><asp:RangeValidator
                                ID="rvQuantityAuthorised" runat="server" ControlToValidate="tbQuantityAuthorised"
                                EnableClientScript="False" ErrorMessage="*" MaximumValue="20000" MinimumValue="0"
                                Type="Integer"></asp:RangeValidator></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="UserID" HeaderText="User ID" SortExpression="UserID" />
                <asp:BoundField DataField="FirstName" HeaderText="First Name" SortExpression="FirstName" />
                <asp:BoundField DataField="LastName" HeaderText="Last Name" SortExpression="LastName" />
                <asp:TemplateField HeaderText="Duration">
                    <ItemTemplate>
                        &nbsp;&nbsp;&nbsp;<asp:TextBox ID="tbDuration" Text="unlimited" Font-Size="XX-Small"
                            Font-Names="Verdana" runat="server" Width="48px" BackColor="LightYellow"></asp:TextBox><asp:RequiredFieldValidator
                                ID="rfvDuration" runat="server" ControlToValidate="tbDuration" EnableClientScript="False"
                                ErrorMessage="*"></asp:RequiredFieldValidator><asp:RangeValidator ID="rvDuration"
                                    runat="server" ControlToValidate="tbDuration" EnableClientScript="False" ErrorMessage="*"
                                    MaximumValue="1000" MinimumValue="0" Type="Integer"></asp:RangeValidator></ItemTemplate>
                    <HeaderTemplate>
                        <strong>Duration<br />
                            &nbsp;&nbsp;&nbsp;in days</strong></HeaderTemplate>
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemTemplate>
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbAuthoriseProduct"
                            runat="server" Width="32px" />
                    </ItemTemplate>
                    <HeaderTemplate>
                        <asp:Button ID="btnSelectAuthoriseProducts" runat="server" Text="select all" OnClick="btnSelectAuthoriseProducts_Click" /><br />
                        Authorise<br />
                        &nbsp;</HeaderTemplate>
                </asp:TemplateField>
            </Columns>
            <EmptyDataTemplate>
                no authorisations pending</EmptyDataTemplate>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
        <asp:Label ID="lblAuthErrorMessage" runat="server" ForeColor="Red"></asp:Label><br />
        <table width="100%">
            <tr>
                <td style="width: 20%">
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnAuthorise" runat="server" Text="authorise" OnClick="btnAuthorise_Click"
                        Width="120px" />
                </td>
                <td style="width: 50%">
                    <asp:RadioButtonList ID="rblAuthoriseAction" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
                        <asp:ListItem Selected="True" Value="leave">Leave items pending if not marked as Authorise</asp:ListItem>
                        <asp:ListItem Value="decline">Decline authorisation for items not marked Authorise</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td style="width: 10%">
                </td>
            </tr>
        </table>
        <br />
        &nbsp;
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAuthoriseOrder" runat="server" Visible="False" Width="100%" Font-Names="Verdana"
        Font-Size="X-Small">
        <table width="100%">
            <tr>
                <td style="width: 80%">
                    <asp:Label ID="Label25aa" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Orders Awaiting Authorisation"></asp:Label>
                </td>
                <td style="width: 20%" align="right">
                    <asp:Button ID="btnAuthOrderGoBack2" runat="server" Text="go back" CausesValidation="false"
                        OnClick="btn_GoToProductListPanel_click" />
                </td>
            </tr>
        </table>
        <br />
        The following order(s) await authorisation.<br />
        <br />
        Select an order to view.<br />
        <br />
        <br />
        <asp:GridView ID="gvAuthoriseOrder" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="False" AllowPaging="True" Width="100%" ShowFooter="True"
            GridLines="None" OnPageIndexChanging="gvAuthoriseOrder_PageIndexChanging">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="btnShowAuthOrder" runat="server" Text="show order" CommandArgument='<%# Container.DataItem("id")%>'
                            OnClick="btnShowAuthOrder_Click" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="id" InsertVisible="False" ReadOnly="True" SortExpression="id"
                    Visible="False" />
                <asp:TemplateField HeaderText="Order Placed">
                    <ItemTemplate>
                        <asp:Label ID="lblAuthCreatedDate" runat="server" Text='<%# Container.DataItem("OrderCreatedDateTime")%>'></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Orderer">
                    <ItemTemplate>
                        <asp:Label ID="lblFirstName" runat="server" Text='<%# Container.DataItem("FirstName")%>'></asp:Label><asp:Label
                            ID="lblLastName" runat="server" Text='<%# Container.DataItem("LastName")%>'></asp:Label>(<asp:Label
                                ID="lblUserID" runat="server" Text='<%# Container.DataItem("UserID")%>'></asp:Label>)</ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="CneeName" HeaderText="Consignee" SortExpression="CneeName" />
            </Columns>
            <EmptyDataTemplate>
                no authorisations pending</EmptyDataTemplate>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
            <PagerStyle HorizontalAlign="Center" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlShowAuthOrder" runat="server" Visible="False" Width="100%" Font-Names="Verdana"
        Font-Size="X-Small">
        <table width="100%">
            <tr>
                <td style="width: 80%">
                    <strong>
                        <asp:Label ID="Label44" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="Authorisation Order Details"></asp:Label></strong>
                </td>
                <td style="width: 20%" align="right">
                    <asp:Button ID="Button1" runat="server" Text="go back" CausesValidation="false" OnClick="btn_GoToProductListPanel_click" />&nbsp;
                </td>
            </tr>
        </table>
        <br />
        <table width="100%">
            <tr>
                <td style="width: 20%">
                    <asp:Label ID="Label27" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Ordered by:"></asp:Label>
                </td>
                <td style="width: 80%">
                    <asp:Label ID="lblAuthOrderOrderedBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="height: 18px">
                    <asp:Label ID="Label26" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Order placed:"></asp:Label>
                </td>
                <td style="height: 18px">
                    <asp:Label ID="lblAuthOrderPlacedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td style="height: 18px">
                    <asp:Label ID="Label40" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Orderer's message:"></asp:Label>
                </td>
                <td style="height: 18px">
                    <asp:Label ID="lblAuthMsgToAuthoriser" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label30" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Font-Bold="true" Text="CONSIGNEE DETAILS"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label28" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Consignee:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderConsignee" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label29" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Attn Of:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderAttnOf" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label31" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Addr 1:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderAddr1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label32" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Addr 2:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderAddr2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label34" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Addr 3:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderAddr3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label33" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Town/City:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label35" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="County/State"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderState" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label36" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Postcode:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label37" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Country:"></asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblAuthOrderCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text=""></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <em><span style="color: #990000">
                        <asp:Label ID="Label66" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text="NOTE: All items in this consignment were available when the order was created. Some items may not now be available if they have been ordered by other users since this order was created."
                            Font-Italic="True"></asp:Label></span></em>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Label ID="Label38" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Font-Bold="true" Text="CONSIGNMENT DETAILS" /><asp:Label ID="Label41" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text=" (items in " /><asp:Label ID="Label42"
                                runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="red" Text="red" /><asp:Label
                                    ID="Label43" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=" require authorisation)" />
                </td>
            </tr>
        </table>
        <br />
        <asp:GridView ID="gvAuthOrderDetails" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="False" Width="100%" GridLines="None">
            <Columns>
                <asp:TemplateField HeaderText="Product Code">
                    <ItemTemplate>
                        <asp:HiddenField ID="hidLogisticProductKey" runat="server" Value='<%# Container.DataItem("LogisticProductKey")%>' />
                        <asp:HiddenField ID="hidQtyAvailable" runat="server" Value='<%# Container.DataItem("QtyAvailable")%>' />
                        <asp:HiddenField ID="hidArchiveFlag" runat="server" Value='<%# Container.DataItem("ArchiveFlag")%>' />
                        <asp:HiddenField ID="hidDeletedFlag" runat="server" Value='<%# Container.DataItem("DeletedFlag")%>' />
                        <asp:Label ID="lblAuthProductCode" runat="server" Text='<%# Container.DataItem("ProductCode")%>'
                            ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Product Date">
                    <ItemTemplate>
                        <asp:Label ID="lblAuthProductDate" runat="server" Text='<%# Container.DataItem("    ")%>'
                            ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Description">
                    <ItemTemplate>
                        <asp:Label ID="lblAuthProductDescription" runat="server" Text='<%# Container.DataItem("ProductDescription")%>'
                            ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Qty">
                    <ItemTemplate>
                        <asp:TextBox ID="tbAuthOrderQty" Width="50px" MaxLength="6" Font-Names="Verdana"
                            Font-Size="xX-Small" BackColor="lightYellow" runat="server" Text='<%# Container.DataItem("ItemsOut")%>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <EmptyDataTemplate>
                no items found in order</EmptyDataTemplate>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
        </asp:GridView>
        <br />
        <asp:Button ID="btnOrderAuthorise" runat="server" Text="grant authorisation" OnClick="btnOrderAuthorise_Click" />
        <asp:Button ID="btnOrderDecline" runat="server" Text="decline authorisation" OnClick="btnOrderDecline_Click" />
        <br />
        <br />
        <asp:Label ID="Label39" runat="server" Text="Message to orderer (optional):"></asp:Label><br />
        <asp:TextBox ID="tbAuthOrderMessage" Width="100%" Font-Names="Verdana" Font-Size="xX-Small"
            BackColor="LightYellow" TextMode="multiLine" runat="server" MaxLength="500"></asp:TextBox><asp:HiddenField
                ID="hidHoldingQueueKey" runat="server"></asp:HiddenField>
    </asp:Panel>
    <asp:Panel ID="pnlMakeAuthorisable" runat="server" Visible="False" Width="100%" Font-Names="Verdana"
        Font-Size="X-Small">
        <table width="100%">
            <tr>
                <td style="width: 80%; height: 27px;">
                    <asp:Label ID="Label45" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Set Product Attribute 'Requires Authorisation'"></asp:Label>
                </td>
                <td style="width: 20%; height: 27px;" align="right">
                    <asp:Button ID="btnRequestAuthGoBack" runat="server" Text="go back" OnClick="btnRequestAuthGoBack_Click" />&nbsp;
                </td>
            </tr>
        </table>
        <br />
        <br />
        <br />
        <table>
            <tr>
                <td style="width: 200px">
                    <asp:Label ID="Label20cx" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Authoriser:"></asp:Label>
                </td>
                <td style="width: 100px">
                    <asp:DropDownList ID="ddlAssignAuthoriser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnSaveAuthorisable1" runat="server" Text="save" OnClick="btnSaveAuthorisable_Click" />
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlRemoveAuthorisable" runat="server" Visible="False" Width="100%"
        Font-Names="Verdana" Font-Size="X-Small">
        <table width="100%">
            <tr>
                <td style="width: 80%; height: 27px;">
                    <asp:Label ID="Label46" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Modify or Remove Product Attribute 'Requires Authorisation'"></asp:Label>
                </td>
                <td style="width: 20%; height: 27px;" align="right">
                    <asp:Button ID="btnRemoveRequestAuthGoBack" runat="server" Text="go back" OnClick="btnRemoveRequestAuthGoBack_Click" />&nbsp;
                </td>
            </tr>
        </table>
        <br />
        EITHER:&nbsp; select another authoriser for this product. NOTE: This will NOT reassign
        waiting authorisations to the new authoriser, and users who currently require authorisation
        will NOT automatically be informed of this change.<br />
        <br />
        <table style="width: 100%">
            <tr>
                <td style="width: 15%" align="right">
                    <asp:Label ID="Label2bb" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Current authoriser:" />
                </td>
                <td style="width: 85%">
                    <asp:Label ID="lblCurrentAuthoriser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="Label18" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="New authoriser:" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlModifyAuthoriser" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlModifyAuthoriser_SelectedIndexChanged" />
                    &nbsp;&nbsp;
                    <asp:CheckBox ID="cbChangeAuthoriserOnAllSelective" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="change on all <i>req auth</i> products <b>where xxx is the current authoriser</b>" />
                    &nbsp;<asp:CheckBox ID="cbChangeAuthoriserOnAll" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="change on all <i>req auth</i> products <b>irrespective of current authoriser</b>"
                        AutoPostBack="True" OnCheckedChanged="cbChangeAuthoriserOnAll_CheckedChanged" />
                </td>
            </tr>
            <tr>
                <td />
                <td />
            </tr>
            <tr>
                <td />
                <td>
                    <asp:Button ID="btnSaveNewAuthoriser" runat="server" Text="save" OnClick="btnSaveNewAuthoriser_Click" />
                </td>
            </tr>
        </table>
        <br />
        <br />
        OR: click the button below to remove &#39;requires authorisation&#39; attribute
        from this product. NOTE: Users who currently require authorisation will NOT automatically
        be informed of this change.<br />
        <br />
        <table style="width: 664px">
            <tr>
                <td style="width: 100px">
                </td>
                <td>
                    <asp:Button ID="btnRemoveAuthorisable" runat="server" Text="remove 'requires authorisation' attribute"
                        OnClick="btnRemoveAuthorisable_Click" />
                </td>
                <td style="width: 100px">
                </td>
                <td style="width: 100px">
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlProductUserProfile" runat="server" Visible="False" Width="100%">
        <table style="font-family: Verdana; font-size: x-small; width: 100%">
            <tr>
                <td style="white-space: nowrap">
                    <asp:Label ID="Label20" runat="server" Visible="True" Font-Bold="true" Text="Set User Permissions for Product " />
                    <asp:Label ID="lblUserPermissionsProductCode" runat="server" ForeColor="Red" Visible="True"
                        Font-Bold="true" Text="" />
                </td>
                <td style="white-space: nowrap" align="right">
                </td>
            </tr>
            <tr>
                <td style="white-space: nowrap">
                    <asp:Button ID="btnShowAllUsers" OnClick="btn_ShowAllUsers_click" runat="server"
                        Text="show all users" />
                    &nbsp;&nbsp;&nbsp;<asp:Label ID="lblLegendSearchForUser" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="search for user:" />&nbsp;<asp:TextBox runat="server"
                            Width="100px" Font-Size="XX-Small" ID="txtProductUserSearch" />
                    &nbsp;<asp:Button ID="btnSearchUsers" OnClick="btn_SearchUsers_click" runat="server"
                        Text="go" />
                    &nbsp;<asp:Label ID="lblLegendNoMatchingRecords" runat="server" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="no matching records"
                        Visible="False" />
                </td>
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblDefaultMaxGrabQty" runat="server" Text="default max grab qty:"
                        Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;<asp:TextBox ID="txtDefaultGrabQty"
                            runat="server" Visible="True" Font-Names="Verdana" Font-Size="XX-Small" Width="50px"
                            Text="0" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnSaveProductUserProfileChanges1"
                        OnClick="btnSaveProductUserProfileChanges_click" runat="server" Text="save changes" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnCancel1" OnClick="btn_GoBackToProductDetail_Click"
                        runat="server" Text="cancel" />
                </td>
            </tr>
        </table>
        <asp:DataGrid ID="grid_ProductUsers" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" AllowSorting="True" OnSortCommand="SortProductUsersGrid" OnPageIndexChanged="grid_ProductUsers_PageChanged"
            AllowPaging="True">
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="Key" HeaderText="Key"></asp:BoundColumn>
                <asp:BoundColumn DataField="UserID" SortExpression="UserID" HeaderText="User ID">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Blue"
                        VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="FirstName" SortExpression="FirstName" HeaderText="First Name">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="LastName" SortExpression="LastName" HeaderText="Last Name">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn HeaderText="Allow Pick">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" Width="6%" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleAllowToPickCheckboxes" OnClick="btnToggleAllowToPickCheckboxes_Click"
                            runat="server" Text="select all" />
                        <br />
                        <asp:Label ID="Label21" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Allow to pick</asp:Label>&nbsp;</HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbAllowToPick" Checked='<%# DataBinder.Eval(Container, "DataItem.AbleToPick") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Apply max grab">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleApplyMaxGrabCheckboxes" OnClick="btnToggleApplyMaxGrabCheckboxes_Click"
                            runat="server" Text="select all" />
                        <br />
                        <asp:Label ID="Label22" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Apply max grab" />&nbsp;</HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbApplyMaxGrab" Checked='<%# DataBinder.Eval(Container, "DataItem.ApplyMaxGrab") %>'
                            runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Qty">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnApplyMaxGrabQty" OnClick="btn_ApplyMaxGrabQty_Click" runat="server"
                            Text="apply default max grab qty" />
                        <a onmouseover="return escape('Copies the value you place in <b>Default max grab quantity</b> to the <b>Max grab quantity</b> for each user - the values are not saved until you click <b>save changes</b>')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <br />
                        <asp:Label ID="Label001" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Max grab qty" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="txtMaxGrabQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="50px" Text='<%# DataBinder.Eval(Container, "DataItem.MaxGrabQty") %>'></asp:TextBox></ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
            <HeaderStyle BorderColor="Gray" Font-Names="Arial" Font-Size="10pt" Wrap="False" />
            <ItemStyle BackColor="WhiteSmoke" Font-Names="Arial" Font-Size="XX-Small" />
            <PagerStyle HorizontalAlign="Center" Mode="NumericPages" />
        </asp:DataGrid>
        <table width="100%" id="tblSaveCancelProductProfile" runat="server" visible="false"
            style="font-family: Verdana; font-size= x-small">
            <tr>
                <td align="left">
                    <asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Users/page:" />
                    <asp:DropDownList ID="ddlUsersPerPage" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlUsersPerPage_SelectedIndexChanged">
                        <asp:ListItem Selected="True">10</asp:ListItem>
                        <asp:ListItem>50</asp:ListItem>
                        <asp:ListItem>200</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;<asp:Label ID="lblLegendSortValueOpenParenthesis" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="(Sorting on " /><asp:Label ID="lblSortValue" runat="server"
                            Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="User ID" /><asp:Label
                                ID="lblLegendSortCloseParenthesis" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                Text=")" />
                </td>
                <td align="right" valign="middle">
                    <asp:Button ID="Button2" OnClick="btnSaveProductUserProfileChanges_click" runat="server"
                        Text="save changes" />
                    &nbsp;&nbsp;<asp:Button ID="Button3" OnClick="btn_GoBackToProductDetail_Click" runat="server"
                        Text="cancel" />
                </td>
            </tr>
        </table>
        <asp:Label ID="lblProductProfileMessage" runat="server" ForeColor="Gray" Font-Size="X-Small"
            Font-Names="Verdana" />
    </asp:Panel>
    <asp:Panel ID="pnlProductPreAuthorise" runat="server" Visible="False" Width="100%">
        <table style="font-family: Verdana; font-size: x-small; width: 100%">
            <tr valign="middle">
                <td valign="middle" style="white-space: nowrap">
                    <asp:Label ID="Label2cc" runat="server" Visible="True" Font-Bold="true" Text="Pre-Authorise Product" />
                    <asp:Label ID="lblPreAuthorisingProduct" runat="server" ForeColor="Red" Visible="True"
                        Font-Bold="true" Text="" />
                </td>
                <td valign="middle" style="white-space: nowrap" align="right">
                </td>
            </tr>
            <tr valign="middle">
                <td valign="middle" style="white-space: nowrap">
                    <asp:Button ID="btnPreAuthoriseShowAllUsers" runat="server" Text="show all users"
                        OnClick="btnPreAuthoriseShowAllUsers_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;Search for user:&nbsp<asp:TextBox runat="server" Width="100px"
                        Font-Size="XX-Small" ID="tbPreAuthoriseUserSearch" />
                    &nbsp;<asp:Button ID="btnPreAuthoriseSearchUsers" runat="server" Text="go" OnClick="btnPreAuthoriseSearchUsers_Click" />
                </td>
                <td valign="middle" style="white-space: nowrap" align="right">
                    <asp:Label ID="Label24" runat="server" Visible="True" Text="Default authorise qty:" />
                    &nbsp;<asp:TextBox ID="tbDefaultPreAuthoriseQty" runat="server" Visible="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="30px" Text="0" />
                    <asp:Label ID="Label25" runat="server" Visible="True" Text="&nbsp;&nbsp;Default duration (days):" />
                    &nbsp;<asp:TextBox ID="tbDefaultDuration" runat="server" Visible="True" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="30px" Text="0" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnSavePreAuthoriseChanges1"
                        OnClick="btnSavePreAuthoriseChanges_Click" runat="server" Text="save changes" />
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnPreAuthoriseGoBackToProductDetails" OnClick="btn_GoBackToProductDetail_Click"
                        runat="server" Text="cancel" />
                </td>
            </tr>
        </table>
        <asp:DataGrid ID="dgPreAuthorise" runat="server" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None"
            ShowFooter="True" AllowSorting="True" OnSortCommand="SortPreAuthoriseGrid" OnItemDataBound="dgPreAuthorise_ItemDataBound">
            <HeaderStyle Font-Size="10pt" Font-Names="Arial" Wrap="False" BorderColor="Gray">
            </HeaderStyle>
            <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
            <ItemStyle Font-Size="XX-Small" Font-Names="Arial" BackColor="WhiteSmoke"></ItemStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="UserProfileKey" HeaderText="Key"></asp:BoundColumn>
                <asp:BoundColumn DataField="UserID" SortExpression="UserID" HeaderText="User ID">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Blue"
                        VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="UserName" SortExpression="UserName" HeaderText="User Name">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:BoundColumn DataField="Department" SortExpression="Department" HeaderText="Dept/CC">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Font-Bold="True" Wrap="False"
                        ForeColor="Blue" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemTemplate>
                        <asp:Label ID="lblCurrentAssignment" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text=""></asp:Label></ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Authorise">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" Width="6%" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnToggleAuthoriseCheckboxes" runat="server" Text="select all" OnClick="btnToggleAuthoriseCheckboxes_Click" />
                        <br />
                        <asp:Label ID="Label21" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Authorise</asp:Label>&nbsp;</HeaderTemplate>
                    <ItemTemplate>
                        <asp:CheckBox ID="cbPreAuthorise" Checked="false" runat="server"></asp:CheckBox>
                    </ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn>
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:HiddenField ID="hidUserProfileKey" Value='<%# Eval("UserProfileKey") %>' runat="server" />
                        <asp:Button ID="btnApplyPreAuthoriseQty" runat="server" Text="apply default auth qty"
                            OnClick="btnApplyPreAuthoriseQty_Click" />
                        <a onmouseover="return escape('Copies the value you place in <b>Default authorise quantity</b> to the <b>Authorise quantity</b> for each user - the values are not saved until you click <b>save changes</b>')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <br />
                        <asp:Label ID="Label0001" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Authorise qty" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="tbPreAuthoriseQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="50px" Text=""></asp:TextBox></ItemTemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn HeaderText="Duration">
                    <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" HorizontalAlign="Center"
                        ForeColor="Gray" VerticalAlign="Bottom"></HeaderStyle>
                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                    <HeaderTemplate>
                        <asp:Button ID="btnApplyDuration" runat="server" Text="apply default duration" OnClick="btnApplyDuration_Click" />
                        <a onmouseover="return escape('Copies the value you place in <b>Default duration</b> to the <b>Duration</b> for each user - the values are not saved until you click <b>save changes</b>')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                        <br />
                        <asp:Label ID="Label000001" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                            Text="Duration" />
                    </HeaderTemplate>
                    <ItemTemplate>
                        <asp:TextBox ID="tbDuration" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="50px" Text=""></asp:TextBox></ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid><asp:Label ID="lblPreAuthoriseMessage" runat="server" ForeColor="Gray"
            Font-Size="X-Small"></asp:Label><br />
        <table style="font-family: Verdana; font-size: x-small; width: 100%">
            <tr valign="middle">
                <td>
                </td>
                <td align="right">
                    <asp:Button ID="btnSavePreAuthoriseChanges2" runat="server" Text="save changes" OnClick="btnSavePreAuthoriseChanges_Click" />
                    &nbsp;&nbsp;<asp:Button ID="btnPreAuthoriseGoBackToProductDetail" OnClick="btn_GoBackToProductDetail_Click"
                        runat="server" Text="cancel" />
                </td>
            </tr>
        </table>
        &nbsp;&nbsp;</asp:Panel>
    <asp:Panel ID="pnlProductGroups" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <strong>
                        <asp:Label ID="Label47" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="Product Groups"></asp:Label></strong>
                </td>
                <td style="width: 50%; height: 26px;" align="right">
                    <asp:Button ID="btnBackFromProductGroupsToList" runat="server" Text="back to list"
                        OnClick="btnBackFromProductGroupsToList_Click" />&nbsp;
                </td>
            </tr>
        </table>
        <br />
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 20%; white-space: nowrap" align="right">
                    <asp:Label ID="Label50" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Product group:"></asp:Label>
                </td>
                <td style="width: 25%; white-space: nowrap">
                    <asp:DropDownList ID="ddlProductGroup" runat="server" Font-Names="Verdana" Width="150px"
                        Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged" />
                </td>
                <td style="width: 50%; white-space: nowrap">
                    <asp:Button ID="btnNewProductGroup" runat="server" Text="new product group" OnClick="btnNewProductGroup_Click" />&nbsp;
                    <asp:Button ID="btnRenameProductGroup" runat="server" Text="rename product group"
                        OnClick="btnRenameProductGroup_Click" />
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
            </tr>
            <tr>
                <td>
                </td>
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label51" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Primary product group owner:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPrimaryProductGroupOwner" runat="server" Width="150px" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                    <asp:Button ID="btnAssignPrimaryProductGroupOwner" runat="server" Text="save" OnClick="btnAssignPrimaryProductGroupOwner_Click" />
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
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label52" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Deputy product group owner:"></asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlDeputyProductGroupOwner" runat="server" Width="150px" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                    <asp:Button ID="btnAssignDeputyProductGroupOwner" runat="server" Text="save" OnClick="btnAssignDeputyProductGroupOwner_Click" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2" style="white-space: nowrap">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td style="white-space: nowrap">
                </td>
                <td>
                    <asp:CheckBox ID="cbUnassignedProductOwnersOnly" runat="server" Text="only show product owners unassigned to a product group"
                        OnCheckedChanged="cbUnassignedProductOwnersOnly_CheckedChanged" AutoPostBack="True"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td style="white-space: nowrap">
                    <asp:Button ID="btnShowProductsInGroup" runat="server" Text="show products in group"
                        Visible="False" OnClick="btnShowProductsInGroup_Click" />
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvProductsInGroup" runat="server" CellPadding="2" Font-Names="Verdana"
                        Font-Size="XX-Small" Visible="False">
                        <EmptyDataTemplate>
                            <asp:Label ID="Label58" runat="server" Text="no products found"></asp:Label></EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlNewProductGroup" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <strong>
                        <asp:Label ID="Label48" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="New Product Group"></asp:Label></strong>
                </td>
                <td style="width: 50%" align="right">
                    <asp:Button ID="btnBackFromNewProductGroupToProductGroup" runat="server" Text="back to product groups"
                        OnClick="btnBackFromNewProductGroupToProductGroup_Click" />&nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label53" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Product group name:"></asp:Label><asp:TextBox ID="tbNewProductGroup" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnCreateNewProductGroup" runat="server" OnClick="btnCreateNewProductGroup_Click"
                        Text="create" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlRenameProductGroup" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <strong>
                        <asp:Label ID="Label49" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="Rename Product Group"></asp:Label></strong>
                </td>
                <td style="width: 50%" align="right">
                    <asp:Button ID="btnBackFromRenameProductGroupToProductGroup" runat="server" Text="back to product groups"
                        OnClick="btnBackFromRenameProductGroupToProductGroup_Click" />&nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label54" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Rename "></asp:Label><asp:Label ID="lblProductGroupToRename" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" ForeColor="Red"></asp:Label><asp:Label
                                ID="Label55" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                                Text=" to "></asp:Label><asp:TextBox ID="tbRenamedProductGroup" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" MaxLength="50" Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnRenameThisProductGroup" runat="server" Text="rename" OnClick="btnRenameThisProductGroup_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlAssociatedProducts" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <strong>
                        <asp:Label ID="Label56" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                            Text="Associated Products for Product "></asp:Label></strong><asp:Label ID="lblAssociatedProductsProductCode"
                                runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                </td>
                <td align="right" style="width: 50%; height: 26px;">
                    <asp:Button ID="btnAssociatedProductsGoBack" runat="server" Text="go back" OnClick="btn_GoBackToProductDetail_Click" />
                </td>
            </tr>
        </table>
        <br />
        <asp:GridView ID="gvAssociatedProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="false" Width="100%" CellPadding="2">
            <Columns>
                <asp:TemplateField ItemStyle-Width="10%">
                    <ItemTemplate>
                        <asp:Button ID="btnRemoveAssociatedProduct" runat="server" CommandArgument='<%# Container.DataItem("LogisticAssociatedProductKey")%>'
                            Text="remove" OnClick="btnRemoveAssociatedProduct_Click" Style="width: 80px" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" SortExpression="ProductCode"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDate" HeaderText="Value Date" SortExpression="ProductDate"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" SortExpression="ProductDescription"
                    ControlStyle-Width="60%" />
                <asp:BoundField DataField="LanguageId" HeaderText="Language" SortExpression="LanguageId"
                    ControlStyle-Width="10%" />
            </Columns>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataTemplate>
                this product has no associated products</EmptyDataTemplate>
        </asp:GridView>
        <br />
        <strong>
            <asp:Label ID="Label57" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                Text="Add Associated Product"></asp:Label></strong><br />
        <br />
        <asp:GridView ID="gvUnassociatedProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            AutoGenerateColumns="false" Width="100%" CellPadding="2">
            <Columns>
                <asp:TemplateField ItemStyle-Width="10%">
                    <ItemTemplate>
                        <asp:Button ID="btnAddAssociatedProduct" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>'
                            Text="add" OnClick="btnAddAssociatedProduct_Click" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" SortExpression="ProductCode"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDate" HeaderText="Value Date" SortExpression="ProductDate"
                    ControlStyle-Width="10%" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" SortExpression="ProductDescription"
                    ControlStyle-Width="60%" />
                <asp:BoundField DataField="LanguageId" HeaderText="Language" SortExpression="LanguageId"
                    ControlStyle-Width="10%" />
            </Columns>
            <RowStyle BackColor="WhiteSmoke" />
            <AlternatingRowStyle BackColor="White" />
            <EmptyDataTemplate>
                no products found</EmptyDataTemplate>
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlProductInactivityAlertStatus" runat="server" Font-Names="Verdana"
        Font-Size="X-Small" Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <asp:Label ID="Label48pias" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Products Using Inactivity Alert" />
                </td>
                <td style="width: 50%; height: 26px;" align="right">
                    <asp:Button ID="btnBackFromProductInactivityAlertStatus1" runat="server" Text="back to product"
                        OnClick="btnBackFromProductInactivityAlertStatus_Click" />
                    &nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:GridView ID="gvProductInactivityAlert" runat="server" CellPadding="2" Font-Names="Verdana"
                        Font-Size="XX-Small" Width="100%">
                        <EmptyDataTemplate>
                            no products are using the inactivity alert feature</EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="height: 26px">
                </td>
                <td style="height: 26px">
                    <asp:Button ID="btnBackFromProductInactivityAlertStatus2" runat="server" Text="back to product"
                        OnClick="btnBackFromProductInactivityAlertStatus_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlConfigureProductInactivityAlert" runat="server" Font-Names="Verdana"
        Font-Size="X-Small" Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%; height: 26px;">
                    <asp:Label ID="Label48cpia" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Configure Product Inactivity Alerts" />
                </td>
                <td style="width: 50%; height: 26px;" align="right">
                    <asp:Button ID="btnBackFromConfigureProductInactivityAlert1" runat="server" Text="back to product"
                        CausesValidation="false" OnClick="btnBackFromConfigureProductInactivityAlert_Click" />
                    &nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label64" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text='NOTE: These buttons view or set Product Inactivity Alerts on <b><font color="red">ALL</font></b> products. A value of 0 disables Product Inactivity Alerts.' />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Button ID="btnViewProductsUsingInactivityAlert" runat="server" Text="view products using inactivity alert"
                        Width="250px" OnClick="btnViewProductsUsingInactivityAlert_Click" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Button ID="btnSetProductInactivityAlertAllExistingProducts" runat="server" Text="set all existing products"
                        Width="250px" OnClick="btnSetProductInactivityAlertAllExistingProducts_Click" />
                    <asp:Label ID="Label53cpia" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="to:" />
                    <asp:TextBox ID="tbProductInactivityAlertPeriodExistingProducts" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="3" Width="50px">0</asp:TextBox><asp:Label ID="Label61"
                            runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="days" />
                    <asp:RequiredFieldValidator ID="rfvProductInactivityAlertPeriodExistingProducts"
                        runat="server" ControlToValidate="tbProductInactivityAlertPeriodExistingProducts"
                        ErrorMessage="required!" Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator><asp:RangeValidator
                            ID="rvProductInactivityAlertPeriodExistingProducts" runat="server" ControlToValidate="tbProductInactivityAlertPeriodExistingProducts"
                            ErrorMessage="must be a number between 0 and 999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="1000" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnSetProductInactivityAlertNewProducts" runat="server" Text="set new products"
                        Width="250px" OnClick="btnSetProductInactivityAlertNewProducts_Click" />
                    <asp:Label ID="Label62" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="to:" />
                    <asp:TextBox ID="tbProductInactivityAlertPeriodNewProducts" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" MaxLength="3" Width="50px">0</asp:TextBox><asp:Label ID="Label63"
                            runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="days" />
                    <asp:RequiredFieldValidator ID="rfvProductInactivityAlertPeriodNewProducts" runat="server"
                        ControlToValidate="tbProductInactivityAlertPeriodNewProducts" ErrorMessage="required!"
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:RequiredFieldValidator><asp:RangeValidator
                            ID="rvProductInactivityAlertPeriodNewProducts" runat="server" ControlToValidate="tbProductInactivityAlertPeriodNewProducts"
                            ErrorMessage="must be a number between 0 and 999" Font-Names="Verdana" Font-Size="XX-Small"
                            MaximumValue="1000" MinimumValue="0" Type="Integer"></asp:RangeValidator><asp:Label
                                ID="lblAvailableToSuperUsersOnly" runat="server" Text="(available to Super Users only)"
                                Visible="False" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    &nbsp; &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnBackFromConfigureProductInactivityAlert2" runat="server" Text="back to product"
                        CausesValidation="false" OnClick="btnBackFromConfigureProductInactivityAlert_Click" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlConfigureCustomLetter" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="Label48ccl" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Configure Custom Letter" />
                </td>
                <td style="width: 50%" align="right">
                    <asp:LinkButton ID="lnkbtnCustomLetterHelp" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        OnClientClick='window.open("help_customletter.pdf", "CustomLetterHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'>custom letter help</asp:LinkButton>&nbsp;
                    &nbsp;<asp:Button ID="btnBackFromConfigureCustomLetter" runat="server" Text="back to product"
                        OnClick="btnBackFromConfigureCustomLetter_Click" />
                    &nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label67" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Custom letter template:" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <FCKeditorV2:FCKeditor ID="fckedCustomLetterTemplate" runat="server" ToolbarSet="CourierSoftware"
                        BasePath="./fckeditor/" Height="300px" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label68" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Custom letter instructions:" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:TextBox ID="tbCustomLetterInstructions" runat="server" Rows="4" TextMode="MultiLine"
                        Width="100%" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnSaveCustomLetterConfiguration" runat="server" OnClick="btnSaveCustomLetterConfiguration_Click"
                        Text="save custom letter configura`tion" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlTemplate" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="LabelScreenTitle" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Screen Title" />
                </td>
                <td style="width: 50%" align="right">
                    <asp:Button ID="btnBackFromWherever" runat="server" Text="back" CausesValidation="false" />
                    &nbsp;
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="Label5Fieldtitle" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Field Title:" />
                    <asp:TextBox ID="tbTextBox" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        MaxLength="50" Width="200px" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:RequiredFieldValidator ID="rfvTextBox" runat="server" ControlToValidate="tbTextBox"
                        ErrorMessage="Please specify whatever it is!" Font-Names="Verdana" Font-Size="XX-Small"
                        InitialValue="0" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Button ID="btnSaveWhatever" runat="server" Text="save" Width="100px" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlProductCreditsControl" runat="server" Visible="false" Width="100%"
        Font-Names="Verdana" Font-Size="X-Small">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small"
                        Text="Product Credits" />
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 20%">
                    <asp:Button ID="btnProductCreditsTemplateReport" runat="server" OnClick="btnProductCreditsTemplateReport_Click"
                        Text="template report" Width="200px" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnProductCreditsRefreshCredits" runat="server" OnClick="btnProductCreditsRefreshCredits_Click"
                        OnClientClick="return confirm(&quot;Refresh credits will apply credits, using the relevant product credit template, to all users who do not have credits in force for a product.\n\nAre you sure you want to do this?&quot;);"
                        Text="refresh credits!" Width="200px" ToolTip="deletes expired for all users, all products - refreshes from templates any product / user that does not have product credit record, with a start date of NOW - you should not normally need to do this as it happens automatically whenever a template is created or modified" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnRefreshAllCredits" runat="server" OnClick="btnResetCredits_Click"
                        Text="recreate all credits!" OnClientClick="return confirm(&quot;This will remove all current credits and recreate them from templates.\n\nAre you ABSOLUTELY sure you want to do this?&quot;);"
                        Width="200px" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnRemoveAllTemplates" runat="server" OnClientClick="return confirm(&quot;This will remove ALL templates!\n\nAre you ABSOLUTELY sure you want to do this?&quot;);" Text="remove all templates!" Width="200px" onclick="btnRemoveAllTemplates_Click" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlConfigureProductCredits" runat="server" Visible="False" Width="100%"
        Font-Names="Verdana" Font-Size="X-Small">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="Label48cclxx" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Configure Product Credits" />&nbsp;<asp:Label ID="lblConfigureProductCreditsProductCode"
                            runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" />
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td>
                    <asp:Label ID="lblConfigureProductCreditsGroups" runat="server" Font-Bold="False"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="Available Groups:" />&nbsp;<asp:DropDownList
                            ID="ddlProductCreditAvailableGroups" runat="server" AutoPostBack="True" Font-Names="Verdana"
                            Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductCreditAvailableGroups_SelectedIndexChanged"
                            Style="height: 18px">
                        </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    <asp:Label ID="lblIncludeGroups" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Included groups:" />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:GridView ID="gvProductCreditsIncludedGroups" runat="server" AutoGenerateColumns="False"
                        CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" >
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnProductCreditsEditEntry" runat="server" CommandArgument='<%# Container.DataItem("id")%>'
                                        OnClick="lnkbtnProductCreditsEditEntry_Click">edit</asp:LinkButton>&nbsp;
                                    <asp:LinkButton ID="lnkbtnProductCreditsApplyUserGroup" runat="server" CommandArgument='<%# Container.DataItem("id")%>'
                                        OnClientClick="return confirm(&quot;This will overwrite all current credits of users in this user group for this product. Are you sure you want to do this?&quot;);"
                                        OnClick="lnkbtnProductCreditsApplyUserGroup_Click">apply</asp:LinkButton>&nbsp;
                                    <asp:LinkButton ID="lnkbtnProductCreditsRemoveEntry" runat="server" CommandArgument='<%# Container.DataItem("id")%>'
                                        OnClick="lnkbtnProductCreditsRemoveEntry_Click" OnClientClick="return confirm(&quot;This will remove the credit template for this user group.\n\nUsers added to the group will not automatically receive default credit settings.\n\nCredits will not be renewed when they expire.\n\nAre you sure you want to do this?&quot;);">remove</asp:LinkButton></ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="GroupName" HeaderText="User Group" ReadOnly="True" SortExpression="GroupName" />
                            <asp:BoundField DataField="Credit" HeaderText="Credit" ReadOnly="True" SortExpression="Credit" >
                            <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Type" SortExpression="EnforceCreditLimit">
                                <ItemTemplate>
                                    <asp:Label ID="lblProductCreditType" runat="server" Text='<%# gvProductCreditType(Container.DataItem) %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="NextRefreshDateTime" HeaderText="Last Changed" ReadOnly="True"
                                SortExpression="NextRefreshDateTime" >
                            <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <asp:BoundField DataField="RefreshMessage" HeaderText="Refresh Interval" ReadOnly="True"
                                SortExpression="RefreshMessage" >
                            <ItemStyle HorizontalAlign="Center" />
                            </asp:BoundField>
                            <%--<asp:BoundField DataField="CarryOverCredit" HeaderText="Carry Over Credit?" ReadOnly="True"
                                SortExpression="CarryOverCredit" />--%>
                            <%--<asp:BoundField DataField="MaxCredits" HeaderText="Max Credits" ReadOnly="True" SortExpression="MaxCredits" />--%>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trEditConfigureProductCredits" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendUserGroup" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="User group:" /><asp:Label ID="lblUserGroup" runat="server"
                            Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" />&nbsp;&nbsp;&nbsp;<asp:Label ID="lblLegendCredit" runat="server" Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Credit:" />
                    &nbsp;<telerik:RadNumericTextBox
                                    ID="rntbCredit" runat="server" DataType="System.Int32" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxValue="999" MinValue="0" ShowSpinButtons="True" Width="55px">
                                    <NumberFormat DecimalDigits="0" ZeroPattern="n" />
                                </telerik:RadNumericTextBox>&nbsp;
                                       <%--<asp:Label ID="lblLegendUserGroup0" runat="server"
                                        Font-Bold="False" Font-Names="Verdana" Font-Size="XX-Small" Text="Next refresh:" />&nbsp;
                                       <telerik:RadDateTimePicker
                                        ID="rdtpNextRefresh" runat="server" Culture="en-GB" DateInput-DateFormat="dd-MMM-yyyy "
                                        DateInput-DisplayDateFormat="dd-MMM-yyyy " FocusedDate="2013-01-01" Font-Names="Verdana"
                                        Font-Size="XX-Small" MinDate="2013-01-01" Width="180px">
                                        <TimeView CellSpacing="-1" Culture="en-GB">
                                        </TimeView>
                                        <TimePopupButton HoverImageUrl="" ImageUrl="" />
                                        <Calendar UseColumnHeadersAsSelectors="False" UseRowHeadersAsSelectors="False" ViewSelectorText="x">
                                        </Calendar>
                                        <DateInput DateFormat="dd-MMM-yyyy hh:mm" DisplayDateFormat="dd-MMM-yyyy hh:mm" LabelWidth="40%">
                                        </DateInput><DatePopupButton HoverImageUrl="" ImageUrl="" />
                                    </telerik:RadDateTimePicker>--%>
                    &nbsp;<asp:Label ID="lblLegendUserGroup1" runat="server" Font-Bold="False" Font-Names="Verdana"
                        Font-Size="XX-Small" Text="Refresh interval:" />&nbsp;<telerik:RadNumericTextBox
                            ID="rntbRefreshInterval" runat="server" DataType="System.Int32" Font-Names="Verdana"
                            Font-Size="XX-Small" MaxValue="999" MinValue="0" ShowSpinButtons="True" Width="55px">
                            <NumberFormat DecimalDigits="0" ZeroPattern="n" />
                        </telerik:RadNumericTextBox><asp:RadioButton ID="rbConfigureProductCreditsIntervalDays"
                            runat="server" Font-Names="Verdana" Font-Size="XX-Small" GroupName="ConfigureProductCreditsInterval"
                            Text="Day(s)" /><asp:RadioButton ID="rbConfigureProductCreditsIntervalMonths" runat="server"
                                Font-Names="Verdana" Font-Size="XX-Small" GroupName="ConfigureProductCreditsInterval"
                                Text="Month(s)" />&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbConfigureProductCreditsEnforce"
                                    runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Enforced" />
                           &nbsp;
                          <%--<asp:CheckBox ID="cbConfigureProductCreditsCarryOverCredit" runat="server"
                          Font-Names="Verdana" Font-Size="XX-Small" Text="Carry over credit" Visible="False" />&nbsp;&nbsp;<asp:Label
                            ID="lblLegendMaxCredits" runat="server" Font-Bold="False" Font-Names="Verdana"
                            Font-Size="XX-Small" Text="Max credits:" Visible="False" /><telerik:RadNumericTextBox ID="rntbMaxCredits"
                                runat="server" DataType="System.Int32" Font-Names="Verdana" Font-Size="XX-Small"
                                MaxValue="999" MinValue="0" ShowSpinButtons="True" Width="40px" Visible="False">
                                <NumberFormat DecimalDigits="0" ZeroPattern="n" />
                            </telerik:RadNumericTextBox>--%>&nbsp;&nbsp;<asp:Button ID="btnConfigureProductCreditsSave"
                                runat="server" OnClick="btnConfigureProductCreditsSave_Click" Text="save" />&nbsp;<asp:Button
                                    ID="btnConfigureProductCreditsSaveAll" runat="server" OnClick="btnConfigureProductCreditsSaveAll_Click"
                                    Text="save to all included groups" />
                    &nbsp;<asp:Button ID="btnConfigureProductCreditsCancel" runat="server" Text="cancel" onclick="btnConfigureProductCreditsCancel_Click" />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnRemoveAllTemplatesForThisProduct" runat="server" onclick="btnRemoveAllTemplatesForThisProduct_Click" Text="remove ALL templates for this product!" Width="300px" />
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnRefreshAllCreditsForThisProduct" runat="server" Text="refresh ALL credits for this product!" Width="300px" onclick="btnRefreshAllCreditsForThisProduct_Click" />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <div>
                        <b><font face="verdana, sans-serif" size="1">
                        <br />
                        Editing, applying and removing product credits</font></b></div>
                    <div>
                    </div>
                    <div>
                        <font face="verdana, sans-serif" size="1">- <b>edit</b> modifies the credit template for this product / user group. The new values are not applied unless either the autorefresh cycle detects that the existing credit record has expired, or you click the 'apply' button.</font></div>
                    <div>
                    </div>
                    <div>
                        <font face="verdana, sans-serif" size="1">- <b>apply</b> removes all credits for this product for the users in this user group and refreshes the credits using the current template values</font></div>
                    <div>
                    </div>
                    <div>
                        <font face="verdana, sans-serif" size="1">- <b>remove</b> removes the template entirely <b><i>AND ALSO</i></b> removes the credits for this product for all users in the user group; the user will revert to Max Grabs for this product</font></div>
                    <br />
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <br />
    <asp:Label ID="lblError" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        ForeColor="Red"></asp:Label><script type="text/javascript">

                                        function OpenHelpWindow(value) {
                                            window.open(value, "Help", "top=10,left=10,width=500,height=400,status=no,toolbar=no,address=no,menubar=no,resizable=no,scrollbars=yes");
                                        }
                                        function ShowImage(value) {
                                            window.open("show_image.aspx?Image=" + value, "ProductImage", "top=10,left=10,width=700,height=700,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=yes");
                                        }
        </script></form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
