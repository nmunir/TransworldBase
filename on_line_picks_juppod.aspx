<%@ Page Language="VB" Theme="AIMSDefault" EnableEventValidation="false" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Net" %>
<script runat="server">
    
    ' CLEANUP NOTES
    
    ' check FormatFinancialServicesProductDisplay
    
    ' CHANGES & REFACTORING REQUIRED
    
    ' clear CustRef fields, etc after placing an order in case another order is placed in the same session
    
    ' VSOE / OECORP
    
    ' handle unlimited multiple orders
    ' improve multiple order interface
    ' improve Calendar Management interface
    ' add Calendar Management unique event number
    ' split code into multiple modules to reduce size
    ' incorporate User Control to handle per customer reference fields
    ' improve CustRef fields (support REGEX)
    ' incorporate images
    ' improve classic view presentation
    ' remove redundant code
    ' put all SQL access into stored procedures
    ' use Massive?
    ' do full app review
    
    ' QUANTUM LEAP
    ' lnkbtnDisplayModeChange
    ' Barcode - Language
    ' Supplier - Misc1
    ' Boxed to ship - Misc2 (Y/N)
    
    Const MODE_JUPITER_STOCK As Int32 = 1
    Const MODE_JUPITER_POD As Int32 = 2
    'Dim gnMode As Int32 = MODE_JUPITER_STOCK
    Dim gnMode As Int32 = MODE_JUPITER_POD

    Const DEVELOPMENT_MACHINE_NAME As String = "CHRISN"
    Const DISPLAY_MODE_CATEGORY As String = "category"
    Const DISPLAY_MODE_ALL As String = "all"
    Const DISPLAY_MODE_SEARCH As String = "search"

    Const CATEGORY_MODE_2_CATEGORIES As Integer = 2
    Const CATEGORY_MODE_3_CATEGORIES As Integer = 3
    
    Const PRODUCT_VIEW_RICH As String = "rich view"
    Const PRODUCT_VIEW_CLASSIC As String = "classic view"
    
    Const MULTIPLE_ADDRESS_PRODUCT_COLUMNS As Int32 = 10
    Const MULTIPLE_ADDRESS_SERVICE_LEVEL_COLUMN As Int32 = MULTIPLE_ADDRESS_PRODUCT_COLUMNS + 1
    Const MULTIPLE_ADDRESS_COST_CENTRE_COLUMN As Int32 = MULTIPLE_ADDRESS_PRODUCT_COLUMNS + 2
    Const MULTIPLE_ADDRESS_ADDRESSEE_COLUMN As Int32 = MULTIPLE_ADDRESS_PRODUCT_COLUMNS + 3
    Const MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT As Int32 = 4

    Const PER_CUSTOMER_CONFIGURATION_NONE As Int32 = 0
    Const PER_CUSTOMER_CONFIGURATION_1_BLACKROCK As Int32 = 1
    Const PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3 As Int32 = 2
    Const PER_CUSTOMER_CONFIGURATION_3_LEGACY_SINGLE_MANDATORY_UNPROMPTED_COST_CENTRE As Int32 = 3
    Const PER_CUSTOMER_CONFIGURATION_4_HYSTER_YALE As Int32 = 4
    Const PER_CUSTOMER_CONFIGURATION_5_KODAK As Int32 = 5
    Const PER_CUSTOMER_CONFIGURATION_6_CIMA As Int32 = 6
    Const PER_CUSTOMER_CONFIGURATION_7_OECORP As Int32 = 7
    Const PER_CUSTOMER_CONFIGURATION_8_MAN As Int32 = 8
    Const PER_CUSTOMER_CONFIGURATION_12_ATKINS As Int32 = 12
    Const PER_CUSTOMER_CONFIGURATION_14_CIPDCOM As Int32 = 14
    Const PER_CUSTOMER_CONFIGURATION_15_ROYLE As Int32 = 15
    Const PER_CUSTOMER_CONFIGURATION_16_VSOE As Int32 = 16
    Const PER_CUSTOMER_CONFIGURATION_17_AAT As Int32 = 17
    Const PER_CUSTOMER_CONFIGURATION_18_PROQUEST As Int32 = 18
    Const PER_CUSTOMER_CONFIGURATION_19_INSIGHT As Int32 = 19
    Const PER_CUSTOMER_CONFIGURATION_20_DAT As Int32 = 20
    Const PER_CUSTOMER_CONFIGURATION_21_UNICRD As Int32 = 21
    Const PER_CUSTOMER_CONFIGURATION_22_RIOTINTO As Int32 = 22
    Const PER_CUSTOMER_CONFIGURATION_23_ARTHRITIS As Int32 = 23
    Const PER_CUSTOMER_CONFIGURATION_24_PROMOVERITAS As Int32 = 24
    Const PER_CUSTOMER_CONFIGURATION_25_CAB As Int32 = 25
    Const PER_CUSTOMER_CONFIGURATION_26_RAMBLERS As Int32 = 26
    Const PER_CUSTOMER_CONFIGURATION_27_OLYMPUS As Int32 = 27
    Const PER_CUSTOMER_CONFIGURATION_28_QUANTUMLEAP As Int32 = 28
    Const PER_CUSTOMER_CONFIGURATION_29_IRWINMITCHELL As Int32 = 29
    Const PER_CUSTOMER_CONFIGURATION_30_JUPITER As Int32 = 30
    
    ' Protected Function ValidBasket
    ' ' HYSTER CC
    ' trPerCustomerConfiguration4Confirmation1
    ' trPerCustomerConfiguration4Confirmation2


    Const CUSTOMER_WURS As Int32 = 579
    Const CUSTOMER_WU As Int32 = 651
    Const CUSTOMER_ACCENTURE As Int32 = 589
    Const CUSTOMER_PROQUEST As Int32 = 148
    Const CUSTOMER_HYSTER As Int32 = 77
    Const CUSTOMER_YALE As Int32 = 680
    Const CUSTOMER_UNICRD As Int32 = 49
    Const CUSTOMER_STRUTT As Int32 = 708
    Const CUSTOMER_ARTHRITIS As Int32 = 711
    Const CUSTOMER_PROMOVERITAS As Int32 = 695
    Const CUSTOMER_CAB As Int32 = 731
    Const CUSTOMER_RAMBLERS As Int32 = 754
    Const CUSTOMER_OLYMPUS As Int32 = 734
    Const CUSTOMER_QUANTUMLEAP As Int32 = 774
    Const CUSTOMER_BOULEVARD As Int32 = 785
    Const CUSTOMER_WUIRE As Int32 = 686
    Const CUSTOMER_IRWINMITCHELL As Int32 = 790
    Const CUSTOMER_JUPITER As Int32 = 784

    Const ACCOUNT_CODE As String = "COURI11111"
    Const LICENSE_KEY As String = "RA61-XZ94-CT55-FH67"

    Const START_ADDRESS_PAGE As Integer = 0
    Const USER_PERMISSION_VIEW_STOCK As Integer = 1024

    Const CALENDAR_MANAGED_NON_BOOKABLE_TOKEN1 As String = "[non-bookable]"
    Const CALENDAR_MANAGED_NON_BOOKABLE_TOKEN2 As String = "virtual"
    
    Const COUNTRY_CODE_UK As Integer = 222
    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    
    Const COUNTRY_CODE_CANADA As Int32 = 38
    Const COUNTRY_CODE_USA As Int32 = 223
    Const COUNTRY_CODE_USA_NYC As Int32 = 256
    
    Const JUPITER_TIMEBAND_FIRSTCHECK_24_START As String = "00.01.00 AM"
    Const JUPITER_TIMEBAND_FIRSTCHECK_24_END As String = "11.58.00 AM"
    Const JUPITER_TIMEBAND_FIRSTCHECK_48_START As String = "00.11.59 AM"
    Const JUPITER_TIMEBAND_FIRSTCHECK_48_END As String = "02.58.00 PM"
    Const JUPITER_TIMEBAND_SECONDCHECK_24_START As String = "12.00.00 PM"
    Const JUPITER_TIMEBAND_SECONDCHECK_24_END As String = "23:59:50 PM"
    Const JUPITER_TIMEBAND_SECONDCHECK_48_START As String = "03.00.00 PM"
    Const JUPITER_TIMEBAND_SECONDCHECK_48_END As String = "23.59.59 PM"

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0
    Private gsUserKey As String, gsCustomerKey As String
    Private gdtBasket As DataTable = New DataTable()
    Private gdtAuthorisationUsage As DataTable
    Private gdtMultiAddressBooking As DataTable
    Private gdtConsignment As DataTable
    Private gdvBasketView As DataView
    Private gdtProductAuthorisationRequired As DataTable
    Private gsProductStatusMessage As String
    Private gbFormatFinancialServicesProductDisplay As Boolean

    Private gdictCurrentMonthBookings As Dictionary(Of Date, String)
    Private gdictSelectionBookings As Dictionary(Of Date, String)
    Private glstSelection As List(Of Date)
    Private gsBasketCountName As String
    
    Enum enumAuthorisationStatus
        NOT_AUTHORISED = -1
        NOT_AUTHORISABLE = 0
        AUTHORISED = 1
    End Enum

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If IsNumeric(Session("UserKey")) Then
            gsUserKey = Session("UserKey")
        Else
            Server.Transfer("session_expired.aspx")
        End If
        If IsNumeric(Session("CustomerKey")) Then
            gsCustomerKey = Session("CustomerKey")
        Else
            Server.Transfer("session_expired.aspx")
        End If

        gsBasketCountName = "SB_BasketItems" & gnMode.ToString

        If Not IsPostBack Then
            Call GetSiteFeatures()
            Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)
    
            If Not IsNothing(Session("SB_ProductSearchCriteria")) Then
                txtProdSearchCriteria.Text = Session("SB_ProductSearchCriteria")
            End If
            
            psVirtualThumbFolder = ConfigLib.GetConfigItem_Virtual_Thumb_URL

            If IsRioTinto() Then  ' Rio Tinto don't want to see category view, but do want categories for searching
                pnlCategorySelection1.Visible = False
                pnlCategorySelection2.Visible = False
            Else
                Call QueryUsesCategories()
            End If

            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_NONE
            Call PerCustomerConfiguration()
            
            If pbUsesCategories Then
                btnShowByCategory.Visible = True
                tblCategorySelection.Visible = True
                Call GetCategories()
                psDisplayMode = DISPLAY_MODE_CATEGORY
                Call ShowCategoriesPanel()
            Else
                btnShowByCategory.Visible = False
                tblCategorySelection.Visible = False
                tblCategorySelection2.Visible = False
            End If

            Call GetConsignorAddress()
            Call GetCountries()
            
            If IsNumeric(Session(gsBasketCountName)) Then
                SetBasketCount(Session(gsBasketCountName))
            Else
                Session(gsBasketCountName) = 0
                SetBasketCount("0")
            End If
    
            chk_QuickMode.Checked = True
            pbInQuickMode = True
    
            Call InitDistributionListDropdown()
            txtProdSearchCriteria.Attributes.Add("onkeypress", "return clickButton(event,'" + btn_SearchProd.ClientID + "')")
            txtSearchCriteriaAddress.Attributes.Add("onkeypress", "return clickButton(event,'" + btn_SearchAddresses.ClientID + "')")

            tbMultipleAddressOrderCustomerRef.Attributes.Add("onkeypress", "return clickButton(event,'" + btnMultipleAddressOrderUpdateOrder.ClientID + "')")
            tbMultipleAddressOrderSpecialInstructions.Attributes.Add("onkeypress", "return clickButton(event,'" + btnMultipleAddressOrderUpdateOrder.ClientID + "')")
            tbMultipleAddressOrderShippingInfo.Attributes.Add("onkeypress", "return clickButton(event,'" + btnMultipleAddressOrderUpdateOrder.ClientID + "')")

            txtCneeName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeAddr1.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeAddr2.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeAddr3.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeCity.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeState.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneePostCode.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeCtcName.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeTel.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCneeEmail.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCustRef1.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCustRef2.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCustRef3.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtCustRef4.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtSpecialInstructions.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            txtShippingInfo.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            
            lnkbtnDisplayModeChange.Text = psProductView
            If IsShowingRichView() Then
                spanQuickModeCheckBox.Visible = True
            Else
                spanQuickModeCheckBox.Visible = False
            End If

            If (Session("UserPermissions") And USER_PERMISSION_VIEW_STOCK) Then
                lblBasketMsg.Visible = False
                lblBasketCount.Visible = False
                lblBasketItemPlural.Visible = False
                btn_viewbasket.Visible = False
                chk_QuickMode.Visible = False
            End If
            Call TrySetDefaultDestinationKey()
            If Not gnMode = MODE_JUPITER_POD Then
                If Request.Cookies("SprintBasket") Is Nothing Then
                    Call CreateSprintBasketCookie()
                Else
                    Dim sCookieBasket = Request.Cookies("SprintBasket")("SB_Basket") & String.Empty
                    Call GetBasketFromSession() ' CN new
                    If (Not sCookieBasket = String.Empty) AndAlso (gdtBasket Is Nothing) AndAlso gnMode = MODE_JUPITER_STOCK Then
                        Call FillBasketFromCookie()
                    End If
                End If
            End If
        End If
        If IsHysterOrYale() Then
            Thread.CurrentThread.CurrentCulture = New CultureInfo("it-IT", False)
        Else
            Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)
        End If
        Response.Buffer = True
        Call SetTitle()
    End Sub
    
    Protected Sub GetBasketFromSession()
        If gnMode = MODE_JUPITER_STOCK Then
            gdtBasket = Session("SB_BasketData")
        Else
            gdtBasket = Session("SB_BasketDataJupiter")
        End If
    End Sub
    
    Protected Sub TrySetDefaultDestinationKey()
        Dim nDefaultDestinationKey As Integer
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("DefaultDestinationGABKey")) Then
                nDefaultDestinationKey = oDataReader("DefaultDestinationGABKey")
            Else
                nDefaultDestinationKey = 0
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
        If nDefaultDestinationKey > 0 Then
            plCneeAddressKey = nDefaultDestinationKey
            Call GetConsigneeAddress()
            Call CaptureBookingInstructions()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Place an Order"
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
        pbSiteShowZeroStockBalances = dr("ShowZeroStock")
        pbSiteApplyMaxGrabs = dr("ApplyMaxGrabs")
        pbMultipleAddressOrders = dr("MultipleAddressOrders")
        pbOrderAuthorisation = dr("OrderAuthorisation")
        pbProductAuthorisation = dr("ProductAuthorisation")
        pbCalendarManagement = dr("CalendarManagement")
        
        pbCustomLetters = dr("CustomLetters")
        
        pbOnDemandProducts = dr("OnDemandProducts")
        pbZeroStockNotifications = dr("Misc3")
        pbShowNotes = dr("ShowNotes")
        pnCategoryMode = dr("CategoryCount")
        If CBool(dr("PostcodeLookup")) Then
            Call EnablePostCodeLookup()
        End If
        
        If CBool(dr("StockOrderCustRef1Visible")) Then
            lblLegendCheckoutCustomerRef1.Text = dr("StockOrderCustRefLabel1Legend") & ":"
            lblLegendConfirmationCustomerRef1.Text = dr("StockOrderCustRefLabel1Legend")
            lblLegendCheckoutCustomerRef1.Visible = True
            lblLegendCheckoutCustomerRef1.Visible = True
            txtCustRef1.Visible = True
            If CBool(dr("StockOrderCustRef1Mandatory")) Then
                lblLegendCheckoutCustomerRef1.ForeColor = Red
                rfvCheckoutCustomerRef1.Enabled = True
                rfvCheckoutCustomerRef1.EnableClientScript = True
            End If
        Else
            lblLegendCheckoutCustomerRef1.Visible = False
            lblLegendConfirmationCustomerRef1.Visible = False
            txtCustRef1.Visible = False
        End If
        If CBool(dr("StockOrderCustRef2Visible")) Then
            lblLegendCheckoutCustomerRef2.Text = dr("StockOrderCustRefLabel2Legend") & ":"
            lblLegendConfirmationCustomerRef2.Text = dr("StockOrderCustRefLabel2Legend")
            lblLegendCheckoutCustomerRef2.Visible = True
            lblLegendCheckoutCustomerRef2.Visible = True
            txtCustRef2.Visible = True
            If CBool(dr("StockOrderCustRef2Mandatory")) Then
                lblLegendCheckoutCustomerRef2.ForeColor = Red
                rfvCheckoutCustomerRef2.Enabled = True
                rfvCheckoutCustomerRef2.EnableClientScript = True
            End If
        Else
            lblLegendCheckoutCustomerRef2.Visible = False
            lblLegendConfirmationCustomerRef2.Visible = False
            txtCustRef2.Visible = False
        End If
        If CBool(dr("StockOrderCustRef3Visible")) Then
            lblLegendCheckoutCustomerRef3.Text = dr("StockOrderCustRefLabel3Legend") & ":"
            lblLegendConfirmationCustomerRef3.Text = dr("StockOrderCustRefLabel3Legend")
            lblLegendCheckoutCustomerRef3.Visible = True
            lblLegendCheckoutCustomerRef3.Visible = True
            txtCustRef3.Visible = True
            If CBool(dr("StockOrderCustRef3Mandatory")) Then
                lblLegendCheckoutCustomerRef3.ForeColor = Red
                rfvCheckoutCustomerRef3.Enabled = True
                rfvCheckoutCustomerRef3.EnableClientScript = True
            End If
        Else
            lblLegendCheckoutCustomerRef3.Visible = False
            lblLegendConfirmationCustomerRef3.Visible = False
            txtCustRef3.Visible = False
        End If
        If CBool(dr("StockOrderCustRef4Visible")) Then
            lblLegendCheckoutCustomerRef4.Text = dr("StockOrderCustRefLabel4Legend") & ":"
            lblLegendConfirmationCustomerRef4.Text = dr("StockOrderCustRefLabel4Legend")
            lblLegendCheckoutCustomerRef4.Visible = True
            lblLegendCheckoutCustomerRef4.Visible = True
            txtCustRef4.Visible = True
            If CBool(dr("StockOrderCustRef4Mandatory")) Then
                lblLegendCheckoutCustomerRef4.ForeColor = Red
                rfvCheckoutCustomerRef4.Enabled = True
                rfvCheckoutCustomerRef4.EnableClientScript = True
            End If
        Else
            lblLegendCheckoutCustomerRef4.Visible = False
            lblLegendConfirmationCustomerRef4.Visible = False
            txtCustRef4.Visible = False
        End If
        lblAuthorisationAdvisory01.Text = dr("AuthorisationAdvisory")

        trMultiAddressOrder.Visible = pbMultipleAddressOrders
    End Sub

    Protected Sub InitDistributionListDropdown()
        Dim lstDistributionListNames As List(Of String) = GetDistributionListNames()
        ddlDistributionList.Items.Add("- please select -")
        For Each s As String In lstDistributionListNames
            ddlDistributionList.Items.Add(s)
        Next
    End Sub
    
    Protected Function IsJupiter() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsJupiter = IIf(gbSiteTypeDefined, gsSiteType = "jupiter", nCustomerKey = CUSTOMER_JUPITER)
    End Function

    Protected Function IsCIMA() As Boolean
        Dim arrCustomerCIMA() As Integer = {44, 51, 490, 515}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsCIMA = IIf(gbSiteTypeDefined, gsSiteType = "cima", Array.IndexOf(arrCustomerCIMA, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsDenton() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsDenton = IIf(gbSiteTypeDefined, gsSiteType = "denton", nCustomerKey = 52)
    End Function
    
    Protected Function IsMLBlackRock() As Boolean
        Dim arrCustomerMerrillLynchBlackRock() As Integer = {12, 23, 30, 46, 558}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsMLBlackRock = IIf(gbSiteTypeDefined, gsSiteType = "blackrock", Array.IndexOf(arrCustomerMerrillLynchBlackRock, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsCustom() As Boolean
        Dim arrCustom() As Integer = {99998, 99999}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsCustom = IIf(gbSiteTypeDefined, gsSiteType = "custom", Array.IndexOf(arrCustom, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsHysterOrYale() As Boolean
        Dim arrHysterOrYale() As Integer = {CUSTOMER_HYSTER, CUSTOMER_YALE}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHysterOrYale = IIf(gbSiteTypeDefined, gsSiteType = "hysteroryale", Array.IndexOf(arrHysterOrYale, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsYale() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsYale = IIf(gbSiteTypeDefined, gsSiteType = "yale", nCustomerKey = CUSTOMER_YALE)
    End Function

    Protected Function IsUNICRD() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsUNICRD = IIf(gbSiteTypeDefined, gsSiteType = "unicrd", nCustomerKey = CUSTOMER_UNICRD)
    End Function

    Protected Function IsBNI() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsBNI = IIf(gbSiteTypeDefined, gsSiteType = "bni", nCustomerKey = 167)
    End Function

    Protected Function IsCIPDCOM() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsCIPDCOM = IIf(gbSiteTypeDefined, gsSiteType = "cipdcom", nCustomerKey = 422)
    End Function
    
    Protected Function IsRoyle() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsRoyle = IIf(gbSiteTypeDefined, gsSiteType = "royle", nCustomerKey = 520)
    End Function
    
    Protected Function IsMan() As Boolean
        Dim arrCustomerMAN() As Integer = {6, 29, 40}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsMan = IIf(gbSiteTypeDefined, gsSiteType = "man", Array.IndexOf(arrCustomerMAN, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsKodak() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsKodak = IIf(gbSiteTypeDefined, gsSiteType = "kodak", nCustomerKey = 352)
    End Function
    
    Protected Function IsOECORP() As Boolean
        Dim arrCustomerOrientExpress() As Integer = {544, 674}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsOECORP = IIf(gbSiteTypeDefined, gsSiteType = "oecorp", Array.IndexOf(arrCustomerOrientExpress, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsVSOE() As Boolean
        Dim arrCustomerVSOE() As Integer = {24, 688, 703}  ' CN add CELLUCOR to VSOE account family
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsVSOE = IIf(gbSiteTypeDefined, gsSiteType = "vsoe", Array.IndexOf(arrCustomerVSOE, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsVSAL() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsVSAL = IIf(gbSiteTypeDefined, gsSiteType = "vsal", nCustomerKey = 8)
    End Function
    
    Protected Function IsRioTinto() As Boolean
        Dim arrCustomerRioTinto() As Integer = {47, 54, 109}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsRioTinto = IIf(gbSiteTypeDefined, gsSiteType = "riotinto", Array.IndexOf(arrCustomerRioTinto, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsAtkins() As Boolean
        ' NB Faithful & Gould 508
        Dim arrCustomerAtkins() As Integer = {151, 368, 417, 418}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsAtkins = IIf(gbSiteTypeDefined, gsSiteType = "atkins", Array.IndexOf(arrCustomerAtkins, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsAAT() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsAAT = IIf(gbSiteTypeDefined, gsSiteType = "aat", nCustomerKey = 654)
    End Function
    
    Protected Function IsAccenture() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsAccenture = IIf(gbSiteTypeDefined, gsSiteType = "accenture", nCustomerKey = CUSTOMER_ACCENTURE)
    End Function
    
    Protected Function IsArthritis() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsArthritis = IIf(gbSiteTypeDefined, gsSiteType = "arthritis", nCustomerKey = CUSTOMER_ARTHRITIS)
    End Function
    
    Protected Function IsPromoVeritas() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsPromoVeritas = IIf(gbSiteTypeDefined, gsSiteType = "promoveritas", nCustomerKey = CUSTOMER_PROMOVERITAS)
    End Function
    
    Protected Function IsCAB() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsCAB = IIf(gbSiteTypeDefined, gsSiteType = "cab", nCustomerKey = CUSTOMER_CAB)
    End Function
    
    Protected Function sSetIsCABVisibility() As String
        If IsCAB() Then
            sSetIsCABVisibility = True
        Else
            sSetIsCABVisibility = False
        End If
    End Function

    Protected Function bSetPackingNoteVisibility() As Boolean
        Return False
        If gnMode = MODE_JUPITER_STOCK Then
            bSetPackingNoteVisibility = True
        Else
            bSetPackingNoteVisibility = False
        End If
    End Function

    Protected Function IsRamblers() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsRamblers = IIf(gbSiteTypeDefined, gsSiteType = "ramblers", nCustomerKey = CUSTOMER_RAMBLERS)
    End Function
    
    Protected Function IsIrwinMitchell() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsIrwinMitchell = IIf(gbSiteTypeDefined, gsSiteType = "irwinmitchell", nCustomerKey = CUSTOMER_IRWINMITCHELL)
    End Function

    Protected Function IsProquest() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsProquest = IIf(gbSiteTypeDefined, gsSiteType = "proquest", nCustomerKey = CUSTOMER_PROQUEST)
    End Function
    
    Protected Function IsInsight() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsInsight = IIf(gbSiteTypeDefined, gsSiteType = "insight", nCustomerKey = 679)
    End Function
    
    Protected Function IsDAT() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsDAT = IIf(gbSiteTypeDefined, gsSiteType = "dat", nCustomerKey = 546)
    End Function
    
    Protected Function IsNotQuantumLeap() As Boolean
        IsNotQuantumLeap = Not IsQuantumLeap()
    End Function
    
    Protected Function IsWURS() As Boolean
        Dim arrCustomerWURS() As Integer = {CUSTOMER_WURS}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWURS = IIf(gbSiteTypeDefined, gsSiteType = "wurs", Array.IndexOf(arrCustomerWURS, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsWU() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWU = IIf(gbSiteTypeDefined, gsSiteType = "wu", nCustomerKey = CUSTOMER_WU)
    End Function
    
    Protected Function IsStrutt() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsStrutt = IIf(gbSiteTypeDefined, gsSiteType = "strutt", nCustomerKey = CUSTOMER_STRUTT)
    End Function
    
    Protected Function IsNotWURS() As Boolean
        IsNotWURS = Not IsWURS()
    End Function

    Protected Function IsNotStruttAndIsNotWURS() As Boolean
        IsNotStruttAndIsNotWURS = Not (IsWURS() Or IsStrutt())
    End Function

    Protected Function IsNotStruttWUHysterYaleCAB() As Boolean
        IsNotStruttWUHysterYaleCAB = Not (IsWURS() Or IsStrutt() Or IsHysterOrYale() Or IsCAB())
    End Function

    Protected Function IsLegacySingleMandatoryCustRef3() As Boolean
        Dim arrCustomerMandatoryBookingRefSites() As Integer = {2, 8, 10, 11, 31, 35, 36, 37, 39, 43, 78, 114, 124, 149, 150, 163, 195, 206, 288, 306, 335, 353, 357, 371, 417, 418}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsLegacySingleMandatoryCustRef3 = IIf(gbSiteTypeDefined, gsSiteType = "mandatorycustref3", Array.IndexOf(arrCustomerMandatoryBookingRefSites, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsLegacySingleMandatoryCostCentre() As Boolean
        Dim arrCustomerMandatoryCostCentreSites() As Integer = {53, 126, 153}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsLegacySingleMandatoryCostCentre = IIf(gbSiteTypeDefined, gsSiteType = "mandatorycostcentre", Array.IndexOf(arrCustomerMandatoryCostCentreSites, nCustomerKey) >= 0)
    End Function
    
    Protected Function IsOlympus() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsOlympus = IIf(gbSiteTypeDefined, gsSiteType = "olympus", nCustomerKey = CUSTOMER_OLYMPUS)
    End Function
    
    Protected Function IsQuantumLeap() As Boolean
        'Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        'IsQuantumLeap = IIf(gbSiteTypeDefined, gsSiteType = "quantumleap", nCustomerKey = CUSTOMER_QUANTUMLEAP)
        Dim arrCustomer() As Integer = {CUSTOMER_QUANTUMLEAP, CUSTOMER_BOULEVARD}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsQuantumLeap = IIf(gbSiteTypeDefined, gsSiteType = "quantum", Array.IndexOf(arrCustomer, nCustomerKey) >= 0)
    End Function

    Protected Function IsWUIRE() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWUIRE = IIf(gbSiteTypeDefined, gsSiteType = "wuire", nCustomerKey = CUSTOMER_WUIRE)
    End Function

    Protected Sub PerCustomerConfiguration()
        If IsMLBlackRock() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_1_BLACKROCK

            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()

            trPerCustomerConfiguration1Checkout1.Visible = True
            trPerCustomerConfiguration1Checkout2.Visible = True
            trPerCustomerConfiguration1Confirmation1.Visible = True
            tbPerCustomerConfiguration1CostCentre.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            Exit Sub
        End If
        
        If IsProquest() Then
            trPerCustomerConfiguration18Checkout1.Visible = True
            Call SetProquestMessage()
            Exit Sub
        End If

        If IsLegacySingleMandatoryCustRef3() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            
            trPerCustomerConfiguration2Checkout1.Visible = True
            trPerCustomerConfiguration2Checkout2.Visible = True
            trPerCustomerConfiguration2Confirmation1.Visible = True
            trPerCustomerConfiguration2Confirmation2.Visible = True
            tbPerCustomerConfiguration2BookingRef.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration2AdditionalRefA.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration2AdditionalRefB.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration2AdditionalRefC.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            If IsVSAL() Then
                trPerCustomerConfiguration18Checkout2.Visible = True
                trPerCustomerConfiguration18Confirmation1.Visible = True
            End If
            Exit Sub
        End If
        
        If IsLegacySingleMandatoryCostCentre() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_3_LEGACY_SINGLE_MANDATORY_UNPROMPTED_COST_CENTRE
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            
            trPerCustomerConfiguration3Checkout1.Visible = True
            trPerCustomerConfiguration3Checkout2.Visible = True
            trPerCustomerConfiguration3Confirmation1.Visible = True
            trPerCustomerConfiguration3Confirmation2.Visible = True
            tbPerCustomerConfiguration3CostCentre.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration3AdditionalRefA.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration3AdditionalRefB.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration3AdditionalRefC.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            Exit Sub
        End If

        If IsCIMA() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_6_CIMA
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration6Checkout1.Visible = True
            trPerCustomerConfiguration6Confirmation1.Visible = True
            trPerCustomerConfiguration6Checkout2.Visible = True
            trPerCustomerConfiguration6Confirmation2.Visible = True
            tbPerCustomerConfiguration6CostCentre.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            
            lblLegendCneeTel.ForeColor = Red
            rfvCneeTel.Enabled = True
            Exit Sub
        End If
        
        If IsAtkins() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_12_ATKINS
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            
            trPerCustomerConfiguration12Checkout1.Visible = True
            trPerCustomerConfiguration12Checkout2.Visible = True
            trPerCustomerConfiguration12Confirmation1.Visible = True
            trPerCustomerConfiguration12Confirmation2.Visible = True
            tbPerCustomerConfiguration12CostCentre.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration12AdditionalRefA.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration12AdditionalRefB.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration12AdditionalRefC.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            Exit Sub
        End If
        
        If IsHysterOrYale() Then
            Call GetCustomerServiceLevels()
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_4_HYSTER_YALE
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration4Checkout1.Visible = True
            trPerCustomerConfiguration4Confirmation1.Visible = True
            trPerCustomerConfiguration4Confirmation3.Visible = True
            Exit Sub
        End If

        If IsKodak() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_5_KODAK
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration5Checkout1.Visible = True
            trPerCustomerConfiguration5Confirmation1.Visible = True
            tbPerCustomerConfiguration5CostCentre.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            Exit Sub
        End If
        
        If IsOECORP() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration2Checkout1.Visible = True
            trPerCustomerConfiguration2Checkout2.Visible = True
            trPerCustomerConfiguration2Confirmation1.Visible = True
            trPerCustomerConfiguration2Confirmation2.Visible = True
            trPerCustomerConfiguration7Checkout1.Visible = True
            trPerCustomerConfiguration7Confirmation1.Visible = True
            'rfvInstructions.Enabled = True
            rfvInstructions.Enabled = False
            lblLegendSpecialInstructionsCheckout.Text = "Instructions:"
            'lblLegendSpecialInstructionsCheckout.ForeColor = Red
            lblLegendSpecialInstructionsConfirmation.Text = "Instructions"
            Exit Sub
        End If
        
        If IsVSOE() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_16_VSOE
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration2Checkout1.Visible = True
            trPerCustomerConfiguration2Checkout2.Visible = True
            trPerCustomerConfiguration2Confirmation1.Visible = True
            trPerCustomerConfiguration2Confirmation2.Visible = True
            rfvInstructions.Enabled = True
            lblLegendSpecialInstructionsCheckout.Text = "Instructions:"
            lblLegendSpecialInstructionsCheckout.ForeColor = Red
            lblLegendSpecialInstructionsConfirmation.Text = "Instructions"
            trPerCustomerConfiguration0CheckoutDeliveryDateCalendar.Visible = True
            Exit Sub
        End If
        
        If IsMan() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration8Checkout1.Visible = True
            trPerCustomerConfiguration8Checkout2.Visible = True
            trPerCustomerConfiguration8Checkout3.Visible = True
            trPerCustomerConfiguration8Confirmation1.Visible = True
            trPerCustomerConfiguration8Confirmation2.Visible = True
            trPerCustomerConfiguration8Confirmation3.Visible = True
            tbPerCustomerConfiguration8BookingRef.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration8PCID.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            tbPerCustomerConfiguration8MDSOrderRef.Attributes.Add("onkeypress", "return clickButton(event,'" + btnConfirmOrder.ClientID + "')")
            Exit Sub
        End If

        If IsCIPDCOM() Then
            lblLegendCheckoutCustomerRef1.Text = "Cost Centre:"
            lblLegendCheckoutCustomerRef1.ForeColor = Red
            rfvCheckoutCustomerRef1.Enabled = True
            rfvCheckoutCustomerRef1.EnableClientScript = True
            lblLegendConfirmationCustomerRef1.Text = "Cost Centre"
            Exit Sub
        End If
        
        If IsRoyle() Then
            lblLegendCheckoutCustomerRef1.Text = "Job Number:"
            lblLegendCheckoutCustomerRef1.ForeColor = Red
            rfvCheckoutCustomerRef1.Enabled = True
            rfvCheckoutCustomerRef1.EnableClientScript = True
            lblLegendConfirmationCustomerRef1.Text = "Job Number"
            Exit Sub
        End If
        
        If IsWURS() Then
            lblLegendSpecialInstructionsCheckout.Visible = False
            txtSpecialInstructions.Visible = False
            lblLegendSpecialInstructionsConfirmation.Visible = False

            If Session("UserType").ToString.ToLower <> "superuser" Then
                trMultiAddressOrder.Visible = False
            Else
                trMultiAddressOrder.Visible = True
            End If
            Exit Sub
        End If
        
        If IsWU() Then
            If Session("UserType").ToString.ToLower <> "superuser" Then
                trMultiAddressOrder.Visible = False
            Else
                trMultiAddressOrder.Visible = True
            End If
            Exit Sub
        End If
        
        If IsAAT() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_17_AAT
            trPerCustomerConfiguration17Checkout1.Visible = True
            ddlPerCustomerConfiguration17ServiceLevel.SelectedIndex = 0
            trPerCustomerConfiguration17Confirmation1.Visible = True
            lblPerCustomerConfiguration17ConfirmationServiceLevel.Text = "Standard Shipping (Courier)"
            Call InitUserCostCentreForAAT()
            If txtShippingInfo.Text = String.Empty Then
                txtShippingInfo.Text = "AAT is a registered charity. No. 1050724"
            End If
            Exit Sub
        End If
        
        If IsInsight() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_19_INSIGHT
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration18Checkout19.Visible = True
            trPerCustomerConfiguration18Confirmation19.Visible = True
            Exit Sub
        End If
        
        If IsDAT() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_20_DAT
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration20Checkout1.Visible = True
            trPerCustomerConfiguration20Confirmation1.Visible = True
            Exit Sub
        End If
        
        If IsDenton() Then
            lblLegendCMCustomerReference.ForeColor = Red
            rfvCMCustomerReference.Enabled = True
            Exit Sub
        End If
        
        If IsAccenture() Then
            lblLegendCMCustomerReference.Text = "WBS:"
            lblLegendCMCustomerReference.ForeColor = Red
            rfvCMCustomerReference.Enabled = True
            tbCMCustomerReference.MaxLength = 8
            Exit Sub
        End If
        
        If IsUNICRD() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_21_UNICRD
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration21Checkout1.Visible = True
            trPerCustomerConfiguration21Confirmation1.Visible = True
            Exit Sub
        End If
        
        If IsRioTinto() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_22_RIOTINTO
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration22Checkout1.Visible = True
            trPerCustomerConfiguration22Confirmation1.Visible = True
            Exit Sub
        End If

        If IsArthritis() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_23_ARTHRITIS
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration23Checkout1.Visible = True
            trPerCustomerConfiguration23Checkout2.Visible = True
            trPerCustomerConfiguration23Confirmation1.Visible = True
            trPerCustomerConfiguration23Confirmation2.Visible = True
            Exit Sub
        End If

        If IsPromoVeritas() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_24_PROMOVERITAS
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration24Checkout1.Visible = True
            trPerCustomerConfiguration24Checkout2.Visible = True
            trPerCustomerConfiguration24Confirmation1.Visible = True
            trPerCustomerConfiguration24Confirmation2.Visible = True
            Exit Sub
        End If
        
        If IsCAB() Then
            Call GetCustomerServiceLevels()
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_25_CAB
            'Call HideStandardCheckoutCustRefArea()
            'Call HideStandardConfirmationCustRefArea()
            trStandardCheckoutCustRef2.Visible = False
            trStandardConfirmationCustRef2.Visible = False

            ' trPerCustomerConfiguration4Checkout1.Visible = True
            trPerCustomerConfiguration25Confirmation1.Visible = True
            Exit Sub
        End If

        If IsRamblers() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_26_RAMBLERS
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration26Checkout1.Visible = True
            trPerCustomerConfiguration26Checkout2.Visible = True
            Exit Sub
        End If
        
        If IsIrwinMitchell() Then
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_29_IRWINMITCHELL
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            trPerCustomerConfiguration29Checkout1.Visible = True
            trPerCustomerConfiguration29Checkout2.Visible = True
            trPerCustomerConfiguration29Confirmation1.Visible = True
            trPerCustomerConfiguration29Confirmation2.Visible = True
            Exit Sub
        End If
        
        If IsQuantumLeap() Then
            lnkbtnDisplayModeChange.Visible = False
            Exit Sub
        End If
        
        If IsWUIRE() Then
            rfvPostCodeLookup.Enabled = False
            lblLegendPostcodeZipcode.ForeColor = lblLegendCneeTel.ForeColor
            Exit Sub
        End If
        
        If gnMode = MODE_JUPITER_POD Then
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_30_JUPITER
            trPerCustomerConfiguration30Checkout1.Visible = True
            trPerCustomerConfiguration30Checkout2.Visible = True
            trPerCustomerConfiguration30Confirmation1.Visible = True
            trPerCustomerConfiguration30Confirmation2.Visible = True
            'Call InitJupiterPrintServiceLevelDropdown()
            lblLegendStockItems.Text = "Printed materials"
            trCheckoutPackingNoteText.Visible = False
            trConfirmationPackingNoteText.Visible = False
        ElseIf IsJupiter() Then
            Call HideStandardCheckoutCustRefArea()
            Call HideStandardConfirmationCustRefArea()
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_30_JUPITER
            trPerCustomerConfiguration30Checkout1.Visible = True
            lblLegendJupiter.Visible = False
            rfvPerCustomerConfiguration30PrintServiceLevel.Visible = False
            ddlPerCustomerConfiguration30PrintServiceLevel.Visible = False
            trPerCustomerConfiguration30Confirmation2.Visible = True
        End If
    End Sub
    
    Protected Sub InitJupiterPrintServiceLevelDropdown()
        Dim bProductQtyExceeds100 As Boolean = False
        For Each dr As DataRow In gdtBasket.Rows
            If dr("QtyToPick") > 100 Then
                bProductQtyExceeds100 = True
                Exit For
            End If
        Next
        ddlPerCustomerConfiguration30PrintServiceLevel.Items.Clear()
        ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("- please select -", 0))
        If isTimeBetween(JUPITER_TIMEBAND_FIRSTCHECK_24_START, JUPITER_TIMEBAND_FIRSTCHECK_24_END, DateTime.Now) Then
            If Not bProductQtyExceeds100 Then
                ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("1 working day turnround", 24))
            End If
            ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("2 working days turnround", 48))
            ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("3 working days turnround", 72))
        ElseIf isTimeBetween(JUPITER_TIMEBAND_FIRSTCHECK_48_START, JUPITER_TIMEBAND_FIRSTCHECK_48_END, DateTime.Now) Then
            ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("2 working days turnround", 48))
            ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("3 working days turnround", 72))
        Else
            ddlPerCustomerConfiguration30PrintServiceLevel.Items.Add(New ListItem("3 working days turnround", 72))
        End If
    End Sub
    
    Protected Function isTimeBetween(ByVal timestart As String, ByVal timeEnd As String, ByVal checkDate As DateTime) As Boolean
        'Dim dtBegin As DateTime = DateTime.Parse(Now.ToShortDateString & " " & timestart)
        'Dim dtEnd As DateTime = DateTime.Parse(Now.ToShortDateString & " " & timeEnd)
        Dim dtBegin As DateTime = DateTime.Parse(Now.ToString("dd-MMM-yyyy") & " " & timestart)
        Dim dtEnd As DateTime = DateTime.Parse(Now.ToString("dd-MMM-yyyy") & " " & timeEnd)
        'If dtBegin > dtEnd Then 'times span midnight
        '    dtEnd = dtEnd.AddDays(1)
        '    If dtBegin > checkDate And dtEnd > checkDate Then 'checkdate is after midnight, make adjustment
        '        dtBegin = dtBegin.AddDays(-1)
        '        dtEnd = dtEnd.AddDays(-1)
        '    End If
        'End If
        If checkDate >= dtBegin AndAlso checkDate < dtEnd Then Return True
        Return False
    End Function

    Protected Sub SetProquestMessage()
        Session("SB_SpecialInstructions") = "SEND ECONOMY"
        txtSpecialInstructions.Text = "SEND ECONOMY"
    End Sub
    
    Protected Sub InitUserCostCentreForAAT()
        Dim sSQL As String = "SELECT Department FROM UserProfile WHERE [key] = " & Session("UserKey")
        Session("SB_BookingRef1") = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0) & String.Empty
    End Sub
    
    Protected Sub HideStandardCheckoutCustRefArea()
        trStandardCheckoutCustRef1.Visible = False
        trStandardCheckoutCustRef2.Visible = False
    End Sub

    Protected Sub HideStandardConfirmationCustRefArea()
        trStandardConfirmationCustRef1.Visible = False
        trStandardConfirmationCustRef2.Visible = False
    End Sub
    
    Sub GetCustomerServiceLevels()
        If IsYale() Then
            ddlPerCustomerConfiguration4Confirmation1ServiceLevel.Items.Clear()
            ddlPerCustomerConfiguration4Confirmation1ServiceLevel.Items.Add(New ListItem("- please select -", 0))
            ddlPerCustomerConfiguration4Confirmation1ServiceLevel.Items.Add(New ListItem("STANDARD", 1))
            ddlPerCustomerConfiguration4Confirmation1ServiceLevel.Items.Add(New ListItem("EXPRESS", 2))
        ElseIf IsCAB() Then
            ddlPerCustomerConfiguration25Confirmation1ServiceLevel.Items.Clear()
            ddlPerCustomerConfiguration25Confirmation1ServiceLevel.Items.Add(New ListItem("- please select -", 0))
            ddlPerCustomerConfiguration25Confirmation1ServiceLevel.Items.Add(New ListItem("STANDARD", 1))
            ddlPerCustomerConfiguration25Confirmation1ServiceLevel.Items.Add(New ListItem("EXPRESS", 2))
        Else
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spStockMngr_TariffAssignment_GetServiceLevelsForCustomer", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
            oParam.Value = CLng(Session("CustomerKey"))
            Try
                oConn.Open()
                ddlPerCustomerConfiguration4Confirmation1ServiceLevel.DataSource = oCmd.ExecuteReader()
                ddlPerCustomerConfiguration4Confirmation1ServiceLevel.DataTextField = "ServiceLevel"
                ddlPerCustomerConfiguration4Confirmation1ServiceLevel.DataValueField = "ServiceLevelKey"
                ddlPerCustomerConfiguration4Confirmation1ServiceLevel.DataBind()
            Catch ex As SqlException
                lblError.Text = ex.Message
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Function GetUsesCategories() As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_UsesCategories_Get", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                If IsDBNull(oDataTable.Rows(0).Item(0)) Then
                    GetUsesCategories = False
                Else
                    GetUsesCategories = oDataTable.Rows(0).Item(0)
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetUsesCategories: " & ex.Message)
        Finally
            oConn.Close()
        End Try
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
            WebMsgBox.Show("SetUsesCategories: " & ex.Message)
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
        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                CountCategories = CInt(oDataTable.Rows(0).Item(0))
            End If
        Catch ex As Exception
            WebMsgBox.Show("CountCategories: " & ex.Message)
        Finally
            oConn.Close()
        End Try
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
        If pbUsesCategories Then
            pnCategoryMode = pnCategoryMode
        End If
    End Sub
    
    Protected Sub EnablePostCodeLookup()
        trPostCodeLookupStartLine.Visible = True
        trPostCodeLookupPreamble.Visible = True
        trPostCodeLookupInput.Visible = True
        trPostCodeLookupOutput.Visible = False
        trPostCodeLookupFinishLine.Visible = True
        tbPostCodeLookup.Attributes.Add("onkeypress", "return clickButton(event,'" + btnFindAddress.ClientID + "')")
    End Sub
    
    Protected Sub ShowProductList()
        If psProductView = PRODUCT_VIEW_RICH Then
            Call ShowClassicProductList()
            Exit Sub
        End If
        Call HideAllPanels()
        pnlProductList.Visible = True
    End Sub
    
    Protected Sub ShowBasket()
        If psProductView = PRODUCT_VIEW_RICH Then
            Call ShowClassicBasket()
            Exit Sub
        End If
        Call HideAllPanels()
        pnlBasket.Visible = True
    End Sub
    
    Protected Sub ShowEmptyBasket()
        Call HideAllPanels()
        pnlEmptyBasket.Visible = True
    End Sub
    
    Protected Sub ShowBookingConfirmationPanel()
        Call HideAllPanels()
        Call ClearSprintBasketCookie()
        pnlBookingConfirmation.Visible = True
    End Sub
    
    Protected Sub ShowAuthorisationPanel()
        Call HideAllPanels()
        pnlRequestAuthorisation.Visible = True
    End Sub
    
    Protected Sub ShowDeliveryAddressPanel()
        Call HideAllPanels()
        pnlDeliveryAddress.Visible = True
    End Sub
    
    Protected Sub ShowSearchAddressListPanel()
        Call HideAllPanels()
        pnlSearchAddress.Visible = True
    End Sub
    
    Protected Sub ShowConfirmBookingPanel()
        Call HideAllPanels()
        pnlConfirmBooking.Visible = True
    End Sub
    
    Protected Sub ShowCompleteBookingPanel()
        Call HideAllPanels()
        pnlConfirmBooking.Visible = True
    End Sub
            
    Sub ShowCompleteMultipleAddressBookingPanel()
        Call HideAllPanels()
        pnlConfirmMultipleAddressBooking.Visible = True
    End Sub
            
    Protected Sub ShowDistributionListPanel()
        Call HideAllPanels()
        pnlDistributionList.Visible = True
    End Sub
            
    Protected Sub ShowBookingQueuedConfirmationPanel()
        Call HideAllPanels()
        pnlBookingQueuedConfirmation.Visible = True
    End Sub

    Protected Sub ShowClassicProductDetail()
        Call HideAllPanels()
        If IsHysterOrYale() Then
            lblLegendUnitValue.Text = "Unit Value (€)"
        End If
        pnlClassicProductDetail.Visible = True
    End Sub

    Protected Sub ShowClassicProductList()
        If psProductView = PRODUCT_VIEW_CLASSIC Then
            Call ShowProductList()
            Exit Sub
        End If
        Call HideAllPanels()
        pnlClassicProductList.Visible = True
    End Sub

    Protected Sub ShowClassicBasket()
        If psProductView = PRODUCT_VIEW_CLASSIC Then
            Call ShowBasket()
            Exit Sub
        End If
        Call HideAllPanels()
        pnlClassicBasket.Visible = True
    End Sub
    
    Protected Sub ShowCustomLetterPanel(ByVal nProductKey As Integer)
        '    Call HideAllPanels()
        '   Call InitCustomLetterFromTemplate(nProductKey)
        '  pnlCustomLetter.Visible = True
    End Sub
    
    Protected Sub ShowFindAvailableProductsPanel()
        Dim sMonthNames() As String = {"-", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
        Dim sSQL As String
        Dim bProductInUse As Boolean
        Dim nProductKey As Integer
        Dim sProductType As String
        Dim lstSelectedDates As New List(Of Date)
        Dim d As New Date
        d = pdtCalendarManagedSelectionX
        Do
            lstSelectedDates.Add(d)
            d = DateAdd(DateInterval.Day, 1, d)
        Loop Until d > pdtCalendarManagedSelectionY
        
        Dim lstAvailableProducts As New List(Of Integer)
        Dim dv As DataView = GetCalendarManagedItemsDataView()
        For Each drv As DataRowView In dv
            nProductKey = drv("ProductKey")
            If Not IsDBNull(drv("LanguageID")) Then
                sProductType = drv("LanguageID").ToString.Trim
            Else
                sProductType = String.Empty
            End If
            If sProductType <> String.Empty Then
                sSQL = "SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & Session("CustomerKey") & " AND ISNULL(LanguageId,'') = '" & sProductType & "' AND CalendarManaged = 1 AND DeletedFlag = 'N' AND ArchiveFlag = 'N' AND LogisticProductKey <> " & nProductKey
                Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
                For Each dr As DataRow In oDataTable.Rows
                    nProductKey = dr("LogisticProductKey")
                    bProductInUse = False
                    For Each dt As Date In lstSelectedDates
                        sSQL = "SELECT 1 FROM CalendarManagedItemDays cmid INNER JOIN CalendarManagedItemEvent cmie ON cmid.EventId = cmie.[id] WHERE LogisticProductKey = " & nProductKey & " AND EventDay = '" & dt.Day & "-" & sMonthNames(dt.Month) & "-" & dt.Year & "' AND ISNULL(cmie.IsDeleted,0) = 0"
                        Dim oDataTable2 As DataTable = ExecuteQueryToDataTable(sSQL)
                        If oDataTable2.Rows.Count > 0 Then
                            bProductInUse = True
                            Exit For
                        End If
                    Next
                    If Not bProductInUse Then
                        lstAvailableProducts.Add(nProductKey)
                    End If
                Next
            End If
        Next
        sSQL = "SELECT LogisticProductKey, ProductCode, ISNULL(ProductDescription,''), ISNULL(LanguageId,'') FROM LogisticProduct WHERE LogisticProductKey IN ("
        For Each sAvailableProduct As String In lstAvailableProducts
            sSQL += sAvailableProduct & ", "
        Next
        sSQL += "0)"
        Dim oDataTable3 As DataTable = ExecuteQueryToDataTable(sSQL)
        gvCMAvailableProducts.DataSource = oDataTable3
        gvCMAvailableProducts.DataBind()
        If oDataTable3.Rows.Count = 0 Then
            btnCMAddAvailableProductsToBasket.Visible = False
            cbCMRemoveNonBookableProductsFromBasket.Visible = False
        Else
            btnCMAddAvailableProductsToBasket.Visible = True
            cbCMRemoveNonBookableProductsFromBasket.Visible = True
        End If
        lblCMAvailableProductsFromDate.Text = pdtCalendarManagedSelectionX.ToString("dd-MMM-yy")
        lblCMAvailableProductsToDate.Text = pdtCalendarManagedSelectionY.ToString("dd-MMM-yy")
        Call HideAllPanels()
        pnlFindAvailableProducts.Visible = True
    End Sub
    
    Protected Sub HideAllPanels()
        pnlCategorySelection1.Visible = False
        pnlCategorySelection2.Visible = False
        pnlProductList.Visible = False
        pnlBasket.Visible = False
        pnlEmptyBasket.Visible = False
        pnlRequestAuthorisation.Visible = False
        pnlDeliveryAddress.Visible = False
        pnlSearchAddress.Visible = False
        pnlConfirmBooking.Visible = False
        pnlBookingConfirmation.Visible = False
        'pnlAddAddressConfirmation.Visible = False
        pnlDistributionList.Visible = False
        pnlConfirmMultipleAddressBooking.Visible = False
        pnlBookingQueuedConfirmation.Visible = False
        
        pnlClassicProductDetail.Visible = False
        pnlClassicBasket.Visible = False
        pnlClassicProductList.Visible = False
        pnlCalendarManaged.Visible = False
        pnlCustomLetter.Visible = False
        pnlFindAvailableProducts.Visible = False
    End Sub
    
    Protected Sub ddlCneeCountry_Changed(ByVal s As Object, ByVal e As EventArgs)
        plCneeCountryKey = CLng(ddlCneeCountry.SelectedItem.Value)
    End Sub
    
    Protected Sub chk_QuickMode_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        If chk_QuickMode.Checked = True Then
            pbInQuickMode = True
        Else
            pbInQuickMode = False
        End If
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
    
    Protected Sub btnShowByCategory_click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtProdSearchCriteria.Text = ""
        psDisplayMode = DISPLAY_MODE_CATEGORY
        gvProductList.PageIndex = 0
        Call DisplayCategories()
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
        Call ShowCategoriesPanel()
    End Sub
    
    Protected Sub ShowCategoriesPanel()
        Call HideAllPanels()
        If pnCategoryMode = 2 Then
            pnlCategorySelection1.Visible = True
        Else
            pnlCategorySelection2.Visible = True
        End If
    End Sub
    
    Protected Sub lnkbtnShowSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psCategory = CStr(e.CommandArgument)
        Repeater2.Visible = True
        Repeater2a.Visible = True
        Repeater3a.Visible = False
        Call GetSubCategories()
    End Sub
    
    Protected Sub lnkbtnShowSubSubCategories_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        psSubCategory = CStr(e.CommandArgument)
        Repeater3a.Visible = True
        Call GetSubSubCategories()
    End Sub
    
    Protected Sub lnkbtnShowProductsByCategory_click(ByVal sender As Object, ByVal e As CommandEventArgs)
        If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
            psSubCategory = CStr(e.CommandArgument)
            lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " "
        Else
            psSubSubCategory = CStr(e.CommandArgument)
            lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " \ " & psSubSubCategory & " "
        End If
        Call BindProductGridDispatcher("ProductCode")
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_ReturnToProducts_click(ByVal s As Object, ByVal e As System.EventArgs)
        If ValidBasket() Then
            Call ShowProductList()
        End If
    End Sub
    
    Protected Sub btn_ContinueWithOrder_click(ByVal s As Object, ByVal e As System.EventArgs)
        Call ShowCategoriesPanel()
    End Sub
        
    Protected Sub btn_GetFromPersonalAddressBook_click(ByVal s As Object, ByVal e As System.EventArgs)
        Call GetFromPersonalAddressBook()
    End Sub
        
    Protected Sub GetFromPersonalAddressBook()
        pbUsingSharedAddressBook = False
        lblLegendAddressBookType.Text = "Personal Address Book"
        lnkbtnUsePersonalAddressBook.Visible = False
        lnkbtnUseSharedAddressbook.Visible = True
        dgAddressBook.CurrentPageIndex = 0
        Call BindAddressBook()
        Call ShowSearchAddressListPanel()
    End Sub
    
    Protected Sub btn_SearchAddresses_Click(ByVal s As Object, ByVal e As System.EventArgs)
        dgAddressBook.CurrentPageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub btn_ShowAllAddresses_click(ByVal s As Object, ByVal e As System.EventArgs)
        txtSearchCriteriaAddress.Text = ""
        ddlAddressFields.SelectedIndex = 0
        dgAddressBook.CurrentPageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub btn_ReturnToConfirmAddressPanel_click(ByVal s As Object, ByVal e As System.EventArgs)
        Call ShowConfirmBookingPanel()
    End Sub
    
    Protected Sub btn_CancelBooking_click(ByVal s As Object, ByVal e As System.EventArgs)
        Call ShowConfirmBookingPanel()
    End Sub
    
    Protected Sub btn_BackToProductList_click(ByVal s As Object, ByVal e As System.EventArgs)
        lblCategoryHeader.Text = ""
        lblError.Text = ""
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_ReturnToMainPanel_click(ByVal s As Object, ByVal e As ImageClickEventArgs)
        lblCategoryHeader.Text = ""
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_ViewCurrentBasket_click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindBasketGrid("ProductCode")
        Call BindAssocProdGrid()
    End Sub
    
    Protected Sub btn_ShowFullProdList_click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblCategoryHeader.Text = ""
        txtProdSearchCriteria.Text = ""
        Session("SB_ProductSearchCriteria") = "_"
        psDisplayMode = DISPLAY_MODE_ALL
        gvProductList.PageIndex = 0
        Call BindProductGridDispatcher("ProductCode")
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_SearchProd_click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblCategoryHeader.Text = ""
        Session("SB_ProductSearchCriteria") = txtProdSearchCriteria.Text
        psDisplayMode = DISPLAY_MODE_SEARCH
        gvProductList.PageIndex = 0
        Call BindProductGridDispatcher("ProductCode")
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_RefreshProdList_click(ByVal sender As Object, ByVal e As EventArgs)
        If psDisplayMode = DISPLAY_MODE_CATEGORY Then
            If pnCategoryMode = CATEGORY_MODE_2_CATEGORIES Then
                lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " "
            Else
                lblCategoryHeader.Text = " Category selection: " & psCategory & " \ " & psSubCategory & " \ " & psSubSubCategory & " "
            End If
        End If
        Call BindProductGridDispatcher("ProductCode")
        Call ShowProductList()
    End Sub
    
    Protected Sub btn_AddToBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsShowingRichView() Then
            Dim hidProductkey As HiddenField = sender.NamingContainer.FindControl("hidProductkey")
            Dim sProductKey As String = hidProductkey.Value.ToString
            AddItemToBasket(sProductKey, bIsFromCookieBasket:=False)
            If Not pbInQuickMode Then
                Call BindBasketGrid("ProductCode")
                Call ShowBasket()
            End If
        Else
            Call AddItemToBasket(0, bIsFromCookieBasket:=False)
            Call BindBasketGrid("ProductCode")
            Call ShowClassicBasket()
        End If
    End Sub
    
    Protected Function BasketContainsCustomLetter() As Boolean
        BasketContainsCustomLetter = False
    End Function
    
    Protected Function ClassicProductGridCustomLetterSelectionCount() As Int32
        Dim nCount As Int32 = 0
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi In dgrdProducts.Items
            cb = CType(dgi.Cells(11).Controls(1), CheckBox)
            If cb.Checked Then
                Dim cellProductKey As TableCell = dgi.Cells(0)    ' Cells(1) is INFO button
                Dim hidCustomLetter As HiddenField = dgi.Cells(1).FindControl("hidClassicCustomLetter1")
                Try
                    Dim bTest As Boolean = CBool(hidCustomLetter.Value)  ' necessary because db default for custom letter attribute is NULL
                Catch ex As Exception
                    hidCustomLetter.Value = False
                End Try
                If CBool(hidCustomLetter.Value) Then
                    nCount += 1
                End If
            End If
        Next
    End Function
    
    Protected Sub btn_AddAssocItemToBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sProductKey As String
        Dim hidProductkey As HiddenField = sender.NamingContainer.FindControl("hidAssProdkey")
        sProductKey = hidProductkey.Value.ToString
        Call AddAssocProdToBasket(sProductKey)
        Call BindBasketGrid("ProductCode")
        Call ShowBasket()
    End Sub
    
    Protected Sub btnSingleAddressConfirmOrder_click(ByVal sender As Object, ByVal e As EventArgs)
        Call PerformSingleAddressOrder()
    End Sub
    
    Protected Function IsStillValidJupiterPrintServiceLevel() As Boolean
        IsStillValidJupiterPrintServiceLevel = True
        Dim nPrintServiceLevel As Int32 = ddlPerCustomerConfiguration30PrintServiceLevel.SelectedValue
        Select Case nPrintServiceLevel
            Case 24
                If isTimeBetween(JUPITER_TIMEBAND_SECONDCHECK_24_START, JUPITER_TIMEBAND_SECONDCHECK_24_END, Today) Then
                    IsStillValidJupiterPrintServiceLevel = False
                End If
            Case 48
                If isTimeBetween(JUPITER_TIMEBAND_SECONDCHECK_48_START, JUPITER_TIMEBAND_SECONDCHECK_48_END, Today) Then
                    IsStillValidJupiterPrintServiceLevel = False
                End If
        End Select
    End Function
    
    Protected Sub PerformSingleAddressOrder()
        If gnMode = MODE_JUPITER_POD Then
            If Not IsStillValidJupiterPrintServiceLevel() Then
                WebMsgBox.Show("Sorry, the Print Service Level you last selected is no longer valid. Please select another Print Service Level.")
                Call InitJupiterPrintServiceLevelDropdown()
                Exit Sub
            End If
        End If
        If Not pbAuthorisationRequired Then
            Call SubmitOrder()
            Call UpdateAuthorisations()
        Else
            Call PlaceOrderOnHold()
            If pbOnDemandProducts AndAlso psOnDemandSessionGUID <> String.Empty Then
                WebMsgBox.Show("ERROR: Your account is incorrectly configured!!! Failure trying to place an order on hold that contains Print On Demand products! This combination of options is not currently supported.")
                Exit Sub
            End If
        End If
        Call ClearOrderSessionVariables()
    End Sub
   
    Protected Sub btnConfirmOrder_click(ByVal s As Object, ByVal e As EventArgs)
        Call PreConfirmOrder()
    End Sub

    Protected Sub PreConfirmOrder()
        If pbCustomLetters AndAlso bBasketContainsCustomLetter() Then
            Call ProcessCustomLetter()
        Else
            Call ConfirmOrder()
        End If
    End Sub
    
    Protected Sub ProcessCustomLetter()
        
    End Sub
    
    Protected Function bBasketContainsCustomLetter() As Boolean   ' CN
        bBasketContainsCustomLetter = False
    End Function
    
    Protected Sub GetAuthoriserForNextDayDelivery()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetAuthoriserForNextDayDelivery", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCustomerKey As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int)
        oParamCustomerKey.Value = Session("CustomerKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                pnAuthoriser = oDataReader.Item(0)
            Else
                If IsVSOE() Then
                    WebMsgBox.Show("Error - no authoriser found to authorise your Express Delivery request. Cannot continue. Please contact your system supervisor. ")
                ElseIf IsProquest() Then
                    WebMsgBox.Show("Error - no authoriser found to authorise your order. Cannot continue. Please contact your system supervisor. ")
                Else
                    WebMsgBox.Show("Error - no authoriser found. Cannot continue. Please contact your account handler. ")
                End If
                Server.Transfer("session_expired.aspx")
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve authoriser. Cannot continue.")
            Server.Transfer("session_expired.aspx")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub ClearCostCalculator()
        ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedIndex = 0
        lblPerCustomerConfiguration4Confirmation2Weight.Text = String.Empty
        lblPerCustomerConfiguration4Confirmation2BasketShippingCost.Text = String.Empty
        hidCostCalculationTrace.Value = String.Empty
    End Sub
    
    Protected Function UserMustAuthorise() As Boolean
        UserMustAuthorise = True
        Dim sSQL As String = String.Empty
        sSQL = "SELECT UserKey FROM LogisticProductAuthoriseExemptions WHERE UserKey = " & Session("UserKey")
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                UserMustAuthorise = False
            End If
        Catch ex As Exception
            WebMsgBox.Show("UserMustAuthorise: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub ConfirmOrder()
        If UserMustAuthorise() Then
            If ((IsVSOE()) AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery01.Checked) Or IsProquest() Then
                pbAuthorisationRequired = True
                Call GetAuthoriserForNextDayDelivery()
            End If
        End If

        If gnMode = MODE_JUPITER_POD Then
            If ddlPerCustomerConfiguration30PrintServiceLevel.SelectedIndex = 0 Then
                WebMsgBox.Show("Please select the required Print Service Level.")
                Exit Sub
            End If
        End If
    
        If IsCAB() Then
            ddlPerCustomerConfiguration25Confirmation1ServiceLevel.SelectedIndex = 0
        End If
        
        If Not pbAuthorisationRequired Then
            trFinalCheckDefault.Visible = True
            trFinalCheckOrderAuthorisation.Visible = False
        Else
            trFinalCheckDefault.Visible = False
            trFinalCheckOrderAuthorisation.Visible = True
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
            If Not IsValidMDSOrderRef() Then
                WebMsgBox.Show("Invalid MDS Order Reference number - this number is already in use")
                Exit Sub
            End If
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_4_HYSTER_YALE Then
            Call ClearCostCalculator()
        End If
        
        Call CaptureBookingInstructions()

        If IsRamblers() Then
            Session("SB_BookingRef3") = ddlRamblersAreaGroup.SelectedItem.Text
        End If
        
        Call SummariseBooking()
        If psRetrievedAddress <> CompressCneeAddress() Then
            lnkbtnSaveAddressInPersonalAddressBook.Visible = True
            If pbAbleToEditGlobalAddressBook Then
                lnkbtnSaveAddressInSharedAddressBook.Visible = True
            End If
        Else
            lnkbtnSaveAddressInPersonalAddressBook.Visible = False
            lnkbtnSaveAddressInSharedAddressBook.Visible = False
        End If
        Call ShowCompleteBookingPanel()
    End Sub
        
    Protected Sub BindProductGridDispatcher(ByVal SortField As String)
        If psDisplayMode <> DISPLAY_MODE_CATEGORY Then
            Call BindProductGrid(SortField, bUseCategories:=False)
        Else
            Call BindProductGrid(SortField, bUseCategories:=True)
        End If
    End Sub
        
    Protected Sub FormatFinancialServicesProductDisplay(ByRef oDataTable As DataTable)
        gbFormatFinancialServicesProductDisplay = True
        oDataTable.Columns.Add(New DataColumn("QtyAvailable", GetType(Long)))
        For Each dr As DataRow In oDataTable.Rows
            Dim nQuantity As Integer
            nQuantity = dr("Quantity")
            dr("QtyAvailable") = nQuantity
            If Not IsDBNull(dr("ApplyMaxGrab")) Then
                If dr("ApplyMaxGrab") = True Then
                    If dr("MaxGrabQty") < nQuantity Then
                        dr("Quantity") = dr("MaxGrabQty")
                    End If
                End If
            End If
        Next
    End Sub
    
    Protected Sub BindProductGrid(ByVal SortField As String, ByVal bUseCategories As Boolean)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim sSearchCriterion As String = Session("SB_ProductSearchCriteria")
        If sSearchCriterion = "" Then
            sSearchCriterion = "_"
        End If
        lblProductMessage.Text = ""
        Dim sProc As String = "spASPNET_Product_GetProducts10"
        If IsJupiter() Then
            If gnMode = MODE_JUPITER_POD Then
                sProc = "spASPNET_Product_GetProductsJupiterPOD"
            Else
                sProc = "spASPNET_Product_GetProductsJupiterStock"
            End If
        End If
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriterion", SqlDbType.NVarChar, 50))
        If bUseCategories Or sSearchCriterion.Trim = String.Empty Then
            sSearchCriterion = "_"
        End If
        oAdapter.SelectCommand.Parameters("@SearchCriterion").Value = sSearchCriterion
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ApplyMaxGrab", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ApplyMaxGrab").Value = IIf(pbSiteApplyMaxGrabs, 1, 0)
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ShowZeroStockBalances", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ShowZeroStockBalances").Value = IIf(pbSiteShowZeroStockBalances, 1, 0)

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GetByCategory", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@GetByCategory").Value = IIf(bUseCategories, 1, 0)

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CategoryMode", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CategoryMode").Value = pnCategoryMode
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@Category").Value = psCategory

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory").Value = psSubCategory

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory2", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SubCategory2").Value = psSubSubCategory

        oAdapter.Fill(oDataTable)
        If IsShowingRichView() Then
            If oDataTable.Rows.Count > 0 Then
                Call FormatFinancialServicesProductDisplay(oDataTable)
                gvProductList.DataSource = oDataTable
                gvProductList.DataBind()
                gvProductList.Visible = True
                btnRefreshProductList.Visible = True
            Else
                If psDisplayMode = DISPLAY_MODE_SEARCH Then
                    lblProductMessage.Text = "No matching products"
                    txtProdSearchCriteria.Focus()
                ElseIf psDisplayMode = DISPLAY_MODE_ALL Then
                    lblProductMessage.Text = "No products available"
                ElseIf psDisplayMode = DISPLAY_MODE_CATEGORY Then
                    lblProductMessage.Text = "No stock available for products with this categorisation"
                Else
                    lblProductMessage.Text = ""
                End If
                gvProductList.Visible = False
                btnRefreshProductList.Visible = False
            End If
        Else
            If oDataTable.Rows.Count > 0 Then
                Dim oDataView As New DataView(oDataTable)
                oDataView.Sort = SortField
                dgrdProducts.DataSource = oDataView
                dgrdProducts.DataBind()
                dgrdProducts.Visible = True
                btnRefreshProductList.Visible = True
            Else
                lblProductMessage.Text = "No products found"
                dgrdProducts.Visible = False
                btnRefreshProductList.Visible = False
            End If
        End If
        oConn.Close()
    End Sub
    
    Protected Sub BindBasketGrid(ByVal SortField As String)
        Call GetBasketFromSession()  ' CN new
        If IsShowingRichView() Then
            If Not IsNothing(gdtBasket) Then
                Call GetBasketFromSession()
                gdvBasketView = New DataView(gdtBasket)
                gdvBasketView.Sort = "OnDemand ASC," & SortField
                If gdvBasketView.Count > 0 Then
                    gvBasket.DataSource = gdvBasketView
                    gvBasket.DataBind()
                    gvBasket.Visible = "True"
                    Call ShowBasket()
                Else
                    Call ShowEmptyBasket()
                End If
            Else
                Call ShowEmptyBasket()
            End If
        Else
            If Not IsNothing(gdtBasket) Then
                lblClassicBasketMessage.Text = ""
                Call GetBasketFromSession()
                gdvBasketView = New DataView(gdtBasket)
                gdvBasketView.Sort = SortField
                If gdvBasketView.Count > 0 Then
                    dgrdBasket.DataSource = gdvBasketView
                    dgrdBasket.DataBind()
                    dgrdBasket.Visible = "True"
                    lnkbtnClassicProceedToCheckout.Visible = True
                    Call ShowClassicBasket()
                Else
                    dgrdBasket.Visible = "False"
                    lblClassicBasketMessage.Text = "There are no items in your basket"
                    lnkbtnClassicProceedToCheckout.Visible = False
                    Call ShowEmptyBasket()
                End If
            Else
                dgrdBasket.Visible = "False"
                lblClassicBasketMessage.Text = "There are no items in your basket"
                lnkbtnClassicProceedToCheckout.Visible = False
                Call ShowEmptyBasket()
            End If
        End If
    End Sub
    
    Protected Sub BindAssocProdGrid()
        'Dim sShowZeroStockBalances As String = ConfigurationManager.AppSettings("ShowZeroStockBalances")
        'Dim sApplyMaxGrabs As String = ConfigurationManager.AppSettings("ApplyMaxGrabs")
        Dim oConn As New SqlConnection(gsConn)
        Dim AssociatedProds As DataTable = New DataTable()
        Dim dr As DataRow
        Dim dgi As DataGridItem
        Dim sProductKey As String
        Dim vw As DataView
        AssociatedProds.Columns.Add(New DataColumn("ProductKey", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("ProductCode", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("ProductDate", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("Description", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("BoxQty", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("UnitWeightGrams", GetType(Double)))
        AssociatedProds.Columns.Add(New DataColumn("UnitValue", GetType(Double)))
        AssociatedProds.Columns.Add(New DataColumn("ProductCategory", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("Subcategory", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("QtyAvailable", GetType(Long)))
        AssociatedProds.Columns.Add(New DataColumn("QtyToPick", GetType(Long)))
        AssociatedProds.Columns.Add(New DataColumn("PDFFileName", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("OriginalImage", GetType(String)))
        AssociatedProds.Columns.Add(New DataColumn("ThumbNailImage", GetType(String)))
        For Each dgi In gvBasket.Items
            Try
                sProductKey = dgi.Cells(0).Text
                Dim oDataSet As New DataSet()
                Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetAssocProds", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ProductKey").Value = CLng(sProductKey)
                oAdapter.Fill(oDataSet, "AssocProducts")
                Dim dtAssProd As DataTable = oDataSet.Tables("AssocProducts")
                Dim drAssProd As DataRow
                For Each drAssProd In dtAssProd.Rows
                    vw = New DataView(AssociatedProds)
                    vw.RowFilter = "ProductKey='" & drAssProd("LogisticProductKey") & "'"
                    If vw.Count = 0 Then
                        dr = AssociatedProds.NewRow()
                        dr("ProductKey") = drAssProd("LogisticProductKey")
                        dr("ProductCode") = drAssProd("ProductCode")
                        dr("ProductDate") = drAssProd("ProductDate")
                        dr("Description") = drAssProd("ProductDescription")
                        dr("BoxQty") = drAssProd("ItemsPerBox")
                        dr("UnitWeightGrams") = drAssProd("UnitWeightGrams")
                        dr("UnitValue") = drAssProd("UnitValue")
                        dr("ProductCategory") = drAssProd("ProductCategory")
                        dr("Subcategory") = drAssProd("Subcategory")
                        dr("QtyAvailable") = drAssProd("Quantity")
                        dr("QtyToPick") = drAssProd("MaxGrab")
                        dr("PDFFileName") = drAssProd("PDFFileName")
                        dr("OriginalImage") = drAssProd("OriginalImage")
                        dr("ThumbNailImage") = drAssProd("ThumbNailImage")
                        AssociatedProds.Rows.Add(dr)
                    End If
                Next drAssProd
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        Next dgi
        If AssociatedProds.Rows.Count > 0 Then
            pnlAssociatedProducts.Visible = True
            gvAssocProducts.DataSource = AssociatedProds
            gvAssocProducts.DataBind()
            gvAssocProducts.Visible = True
        Else
            gvAssocProducts.Visible = False
        End If
    End Sub
    
    Protected Sub GetAddressBookPermissions()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_UserProfile_GetProfileFromKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
        Try
            oAdapter.Fill(oDataTable)
            Dim dr As DataRow = oDataTable.Rows(0)
            pbAbleToViewGlobalAddressBook = dr("AbleToViewGlobalAddressBook")
            pbAbleToEditGlobalAddressBook = dr("AbleToEditGlobalAddressBook")
        Catch ex As Exception
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        lnkbtnGetFromSharedAddressBook.Visible = pbAbleToViewGlobalAddressBook
        lnkbtnSaveAddressInSharedAddressBook.Visible = pbAbleToEditGlobalAddressBook
    End Sub
    
    Protected Sub BindAddressBook()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddresses", oConn)
        Dim sSearchCriteria As String = txtSearchCriteriaAddress.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        lblAddressMessage.Text = ""
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UseSharedAddressBook", SqlDbType.Bit))
            oAdapter.SelectCommand.Parameters("@UseSharedAddressBook").Value = pbUsingSharedAddressBook
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FieldMask", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@FieldMask").Value = ddlAddressFields.SelectedValue  ' 0=all fields, 1=Company Name
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            If IsNumeric(Session("UserKey")) Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            Else
                Server.Transfer("error.aspx")
            End If

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

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
                lblAddressMessage.Text = "No addresses found. Please refine your search and try again." ' INAPPROPRIATE MESSAGE IF NOT SEARCHING !!
            End If
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
        
    Protected Sub gvBasket_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        Dim sProductKey As String
        Dim cell_ProductKey As TableCell = e.Item.Cells(0)
        If IsNumeric(cell_ProductKey.Text) Then
            sProductKey = cell_ProductKey.Text
            If e.CommandSource.CommandName = "Remove" Then
                RemoveItemFromBasket(sProductKey)
                BindBasketGrid("ProductCode")
            End If
        End If
    End Sub
    
    Protected Sub dgAddressBook_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "select" Then
            Dim cell_Address As TableCell = e.Item.Cells(1)
            If IsNumeric(cell_Address.Text) Then
                plCneeAddressKey = CLng(cell_Address.Text)
                ResetFields()
                GetConsigneeAddress()
                ShowDeliveryAddressPanel()
            End If
        End If
    End Sub
        
    Protected Sub gvProductList_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hlnk_PDF As HyperLink = CType(gvr.Cells(2).FindControl("hlnk_PDF"), HyperLink)
            Dim sProductKey As String = DataBinder.Eval(e.Row.DataItem, "LogisticProductKey").ToString
            Try
                If File.Exists(MapPath(ConfigLib.GetConfigItem_Virtual_PDF_URL & sProductKey & ".pdf")) Then
                    hlnk_PDF.Visible = True
                    hlnk_PDF.NavigateUrl = ConfigLib.GetConfigItem_Virtual_PDF_URL & sProductKey & ".pdf"
                Else
                    hlnk_PDF.Visible = False
                End If
            Catch ex As Exception
                hlnk_PDF.Visible = False
            End Try

            Dim hidNotes As HiddenField = CType(e.Row.FindControl("hidNotes"), HiddenField)
            Dim lbl As Label = CType(e.Row.FindControl("lblProductNotes"), Label)
            Dim tbl As HtmlTable = CType(e.Row.FindControl("tblProductNotes"), HtmlTable)
            If hidNotes.Value = String.Empty OrElse Not pbShowNotes Then
                tbl.Visible = False
            End If

            Dim lblCostCentreLegend As Label = CType(gvr.Cells(2).FindControl("lblCostCentreLegend"), Label)
            Dim lblCostCentre As Label = CType(gvr.Cells(2).FindControl("lblCostCentre"), Label)
            If lblCostCentre.Text = String.Empty Then
                lblCostCentreLegend.Visible = False
            End If

        End If
    End Sub
    
    Protected Sub gvAssocProducts_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim hlnk_AssocPDF As HyperLink = CType(e.Row.FindControl("hlnk_AssocPDF"), HyperLink)
            Dim sProductKey As String = DataBinder.Eval(e.Row.DataItem, "ProductKey").ToString
            Try
                If File.Exists(MapPath(ConfigLib.GetConfigItem_Virtual_PDF_URL & sProductKey & ".pdf")) Then
                    hlnk_AssocPDF.Visible = True
                    hlnk_AssocPDF.NavigateUrl = ConfigLib.GetConfigItem_Virtual_PDF_URL & sProductKey & ".pdf"
                Else
                    hlnk_AssocPDF.Visible = False
                End If
            Catch
                hlnk_AssocPDF.Visible = False
            End Try
        End If
    End Sub
        
    Protected Sub dgAddressBook_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        Dim dgpcea As DataGridPageChangedEventArgs = e
        'dgAddressBook.CurrentPageIndex = e.NewPageIndex
        pnAddressPage = dgpcea.NewPageIndex
        dgAddressBook.CurrentPageIndex = pnAddressPage
        Call BindAddressBook()
    End Sub

    Protected Sub gvProductList_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvProductList.PageIndex = e.NewPageIndex
        Call BindProductGridDispatcher("ProductCode")
    End Sub

    Protected Sub SortBasketColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Dim sSortExpression = E.SortExpression
        If sSortExpression = "Language" Then
            sSortExpression = "LanguageId"
        End If
        Call BindBasketGrid(sSortExpression)
    End Sub
    
    Protected Sub SortProductColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Call BindProductGridDispatcher(E.SortExpression)
    End Sub
    
    Protected Sub GetCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim sProc As String = "spASPNET_Product_GetCategoriesForUser"
        If IsJupiter() Then
            If gnMode = MODE_JUPITER_POD Then
                sProc = "spASPNET_Product_GetCategoriesForUserJupiterPOD"
            Else
                sProc = "spASPNET_Product_GetCategoriesForUserJupiterStock"
            End If
        End If
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
        lblError.Text = ""
        Try
            oAdapter.Fill(oDataSet, "Categories")
            Dim iCount As Integer = oDataSet.Tables(0).Rows.Count
            If iCount > 0 Then
                If pnCategoryMode <> CATEGORY_MODE_3_CATEGORIES Then
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
                pnCategoryMode = CATEGORY_MODE_2_CATEGORIES
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
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
    
    Protected Function GetOnDemandTemplate(ByVal sLogisticProductKey As String) As String
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT * FROM LogisticProductOnDemandTemplate WHERE LogisticProductKey = " & sLogisticProductKey)
        If oDataTable.Rows.Count = 1 Then
            GetOnDemandTemplate = oDataTable.Rows(0).Item("OnDemandTemplate")
        Else
            GetOnDemandTemplate = String.Empty
        End If
    End Function
    
    Protected Sub AddItemToBasket(ByVal sProductKey As String, ByVal bIsFromCookieBasket As Boolean)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If IsShowingRichView() Or bIsFromCookieBasket Then
            Dim dr As DataRow
            Dim oDataReader As SqlDataReader = Nothing
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetFromKey7", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int)
            oParamProductKey.Value = CLng(sProductKey)
            Dim oParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
            oParamUserKey.Value = Session("UserKey")
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                Call CreateBasketIfNull()
                gdvBasketView = New DataView(gdtBasket)
                gdvBasketView.RowFilter = "ProductKey='" & sProductKey & "'"
                If gdvBasketView.Count = 0 Then
                    dr = gdtBasket.NewRow()
                    dr("ProductKey") = sProductKey
                    dr("ProductCode") = oDataReader("ProductCode") & ""
                    dr("ProductDate") = oDataReader("ProductDate") & ""
                    dr("Description") = oDataReader("ProductDescription") & ""
                    dr("LanguageID") = oDataReader("LanguageID") & ""
                    dr("BoxQty") = oDataReader("ItemsPerBox") & ""
                    If Not IsDBNull(oDataReader("UnitWeightGrams")) AndAlso IsNumeric(oDataReader("UnitWeightGrams")) Then
                        dr("UnitWeightGrams") = oDataReader("UnitWeightGrams")
                    Else
                        dr("UnitWeightGrams") = 0
                    End If
                    If Not IsDBNull(oDataReader("UnitValue")) AndAlso IsNumeric(oDataReader("UnitValue")) Then
                        dr("UnitValue") = oDataReader("UnitValue")
                    Else
                        dr("UnitValue") = 0
                    End If
                    If Not IsDBNull(oDataReader("UnitValue2")) AndAlso IsNumeric(oDataReader("UnitValue2")) Then
                        dr("UnitValue2") = oDataReader("UnitValue2")
                    Else
                        dr("UnitValue2") = 0
                    End If
                    If Not IsDBNull(oDataReader("Quantity")) AndAlso IsNumeric(oDataReader("Quantity")) Then
                        dr("QtyAvailable") = oDataReader("Quantity")
                    Else
                        dr("QtyAvailable") = 0
                    End If
                    dr("QtyToPick") = 1
                    dr("PDFFileName") = oDataReader("PDFFileName") & ""
                    dr("OriginalImage") = oDataReader("OriginalImage") & ""
                    dr("ThumbNailImage") = oDataReader("ThumbNailImage") & ""
                    dr("Notes") = oDataReader("Notes") & ""
                    If Not IsDBNull(oDataReader("CalendarManaged")) Then
                        dr("CalendarManaged") = oDataReader("CalendarManaged")
                    Else
                        dr("CalendarManaged") = False
                    End If
                    dr("OnDemand") = 0
                    dr("OnDemandPriceList") = False
                    If Not IsDBNull(oDataReader("CustomLetter")) Then
                        dr("CustomLetter") = oDataReader("CustomLetter")
                    Else
                        dr("CustomLetter") = False
                    End If
                    gdtBasket.Rows.Add(dr)
                    Session(gsBasketCountName) = Session(gsBasketCountName) + 1
                    SetBasketCount(Session(gsBasketCountName))
                    Call SaveBasketToSession()
                    gdvBasketView.RowFilter = ""
                End If
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                If Not oDataReader Is Nothing Then
                    oDataReader.Close()
                End If
                oConn.Close()
            End Try
        Else
            Dim dgi As DataGridItem
            Dim cb As CheckBox
            For Each dgi In dgrdProducts.Items
                cb = CType(dgi.Cells(11).Controls(1), CheckBox)
                If cb.Checked Then
                    Dim dr As DataRow
                    Dim cellProductKey As TableCell = dgi.Cells(0)
                    ' Cells(1) is INFO button
                    Dim hidCalendarManaged As HiddenField = dgi.Cells(1).FindControl("hidClassicCalendarManaged1")
                    Try
                        Dim bTest As Boolean = CBool(hidCalendarManaged.Value)  ' necessary because db default for calendar managed attribute is NULL
                    Catch ex As Exception
                        hidCalendarManaged.Value = False
                    End Try
                    Dim hidCustomLetter As HiddenField = dgi.Cells(1).FindControl("hidClassicCustomLetter1")
                    Try
                        Dim bTest As Boolean = CBool(hidCustomLetter.Value)  ' necessary because db default for custom letter attribute is NULL
                    Catch ex As Exception
                        hidCustomLetter.Value = False
                    End Try
                    Dim hidOnDemand As HiddenField = dgi.Cells(1).FindControl("hidOnDemand1")
                    Try
                        Dim bTest As Boolean = CBool(hidOnDemand.Value)  ' necessary because db default for OnDemand attribute is NULL
                    Catch ex As Exception
                        hidOnDemand.Value = 0
                    End Try
                    Dim hidOnDemandPriceList As HiddenField = dgi.Cells(1).FindControl("hidOnDemandPriceList1")
                    Try
                        Dim bTest As Boolean = CBool(hidOnDemandPriceList.Value)  ' necessary because db default for OnDemand attribute is NULL
                    Catch ex As Exception
                        hidOnDemandPriceList.Value = 0
                    End Try
                    Dim cellProductCode As TableCell = dgi.Cells(2)
                    Dim cellProductDate As TableCell = dgi.Cells(3)
                    Dim cellDescription As TableCell = dgi.Cells(4)
                    ' Cells(5) is Dept Id
                    Dim cellLanguage As TableCell = dgi.Cells(6)
                    Dim cellUnitValue As TableCell = dgi.Cells(7)
                    Dim cellBoxQty As TableCell = dgi.Cells(8)
                    Dim cellUnitWeightGrams As TableCell = dgi.Cells(9)
                    Dim cellQtyAvailable As TableCell = dgi.Cells(10)

                    Dim sCellProductKey As String = cellProductKey.Text
                    Dim sProductCode As String = cellProductCode.Text
                    Dim sProductDate As String = cellProductDate.Text
                    Dim sDescription As String = cellDescription.Text
                    Dim sLanguage As String = cellLanguage.Text
                    Dim sBoxQty As String = cellBoxQty.Text ' hidden
                    sBoxQty = "0" ' hidden
                    Dim lQtyAvailable As Long = CLng(cellQtyAvailable.Text)
                    ' Dim dblUnitValue As Double = CDbl(cellUnitValue.Text) ' hidden
                    Dim dblUnitValue As Double = CDbl(cellUnitValue.Text.Substring(1))
                    Dim dblUnitWeightGrams As Double
                    If IsNumeric(cellUnitWeightGrams.Text) Then          ' this is a 'workaround'. Hmmm.
                        dblUnitWeightGrams = CDbl(cellUnitWeightGrams.Text) ' hidden
                    Else
                        dblUnitWeightGrams = 0
                    End If
                    ' Dim dblUnitWeightGrams = 0 ' hidden
                    Call CreateBasketIfNull()
                    Call GetBasketFromSession()
                    gdvBasketView = New DataView(gdtBasket)
                    gdvBasketView.RowFilter = "ProductKey='" & sCellProductKey & "'"
                    If gdvBasketView.Count = 0 Then
                        dr = gdtBasket.NewRow()
                        dr("ProductKey") = sCellProductKey
                        dr("ProductCode") = sProductCode
                        dr("ProductDate") = sProductDate
                        dr("Description") = sDescription
                        dr("LanguageID") = sLanguage
                        dr("BoxQty") = sBoxQty
                        dr("UnitWeightGrams") = dblUnitWeightGrams
                        dr("UnitValue") = dblUnitValue
                        dr("UnitValue2") = 0 ' not available in classic view
                        dr("QtyAvailable") = lQtyAvailable
                        dr("PDFFileName") = String.Empty ' not available in classic view
                        dr("OriginalImage") = String.Empty ' not available in classic view
                        dr("ThumbNailImage") = String.Empty ' not available in classic view
                        dr("Notes") = String.Empty ' not available in classic view
                        dr("CalendarManaged") = hidCalendarManaged.Value
                        dr("OnDemand") = hidOnDemand.Value
                        dr("OnDemandPriceList") = hidOnDemandPriceList.Value
                        dr("CustomLetter") = hidCustomLetter.Value
                        gdtBasket.Rows.Add(dr)
                        Session(gsBasketCountName) = Session(gsBasketCountName) + 1
                        SetBasketCount(Session(gsBasketCountName))
                        Call SaveBasketToSession()
                    End If
                    gdvBasketView.RowFilter = ""
                End If
            Next dgi
            Call ShowClassicBasket()
        End If
    End Sub
    
    Protected Sub CreateBasketIfNull()
        Call GetBasketFromSession()
        If IsNothing(gdtBasket) Then
            gdtBasket = New DataTable()
            gdtBasket.Columns.Add(New DataColumn("ProductKey", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("ProductCode", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("ProductDate", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("Description", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("LanguageID", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("BoxQty", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("UnitWeightGrams", GetType(Double)))
            gdtBasket.Columns.Add(New DataColumn("UnitValue", GetType(Double)))
            gdtBasket.Columns.Add(New DataColumn("UnitValue2", GetType(Double)))
            gdtBasket.Columns.Add(New DataColumn("QtyAvailable", GetType(Long)))
            gdtBasket.Columns.Add(New DataColumn("QtyToPick", GetType(Long)))
            gdtBasket.Columns.Add(New DataColumn("PDFFileName", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("OriginalImage", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("ThumbNailImage", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("Notes", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("Authorised", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("CalendarManaged", GetType(Boolean)))
            gdtBasket.Columns.Add(New DataColumn("OnDemand", GetType(Integer)))
            gdtBasket.Columns.Add(New DataColumn("OnDemandPriceList", GetType(Integer)))
            gdtBasket.Columns.Add(New DataColumn("OnDemandTemplate", GetType(String)))
            gdtBasket.Columns.Add(New DataColumn("CustomLetter", GetType(Boolean)))
            Call SaveBasketToSession()
        End If
    End Sub
        
    Protected Sub AddAssocProdToBasket(ByVal sProductKey As String)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        
        Dim dr As DataRow
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetFromKey7", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int)
        oParamProductKey.Value = CLng(sProductKey)
        Dim oParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParamUserKey.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            Call CreateBasketIfNull() ' probably don't need this since the basket must exist to get here
            Call GetBasketFromSession()
            gdvBasketView = New DataView(gdtBasket)
            gdvBasketView.RowFilter = "ProductKey='" & sProductKey & "'"
            If gdvBasketView.Count = 0 Then
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
                If Not IsDBNull(oDataReader("LanguageID")) Then
                    dr("LanguageID") = oDataReader("LanguageID")
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
                dr("QtyToPick") = 1
                If Not IsDBNull(oDataReader("PDFFileName")) Then
                    dr("PDFFileName") = oDataReader("PDFFileName")
                End If
                If Not IsDBNull(oDataReader("OriginalImage")) Then
                    dr("OriginalImage") = oDataReader("OriginalImage")
                End If
                If Not IsDBNull(oDataReader("ThumbNailImage")) Then
                    dr("ThumbNailImage") = oDataReader("ThumbNailImage")
                End If
                If Not IsDBNull(oDataReader("Notes")) Then
                    dr("Notes") = oDataReader("Notes")
                Else
                    dr("Notes") = String.Empty
                End If
                If Not IsDBNull(oDataReader("CalendarManaged")) Then
                    dr("CalendarManaged") = oDataReader("CalendarManaged")
                Else
                    dr("CalendarManaged") = False
                End If
                dr("OnDemand") = 0
                dr("OnDemandPriceList") = 0
                dr("CustomLetter") = False
                gdtBasket.Rows.Add(dr)
                Session(gsBasketCountName) = Session(gsBasketCountName) + 1
                SetBasketCount(Session(gsBasketCountName))
                Call SaveBasketToSession()
                gdvBasketView.RowFilter = ""
            End If
        Catch ex As SqlException
            'Server.Transfer("error.aspx")
            lblError.Text = ex.ToString
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
            End If
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SetBasketCount(ByVal sCount As String)
        Dim nCount As Integer
        nCount = CInt(sCount)
        If nCount = 1 Then
            lblBasketItemPlural.Text = "item"
        Else
            lblBasketItemPlural.Text = "items"
        End If
        lblBasketCount.Text = sCount
    End Sub
    
    Protected Sub RemoveItemFromBasket(ByVal sProductKey As String)
        lblError.Text = ""
        lblBasketMessage1.Text = ""
        lblBasketMessage2.Text = ""
        
        Call GetBasketFromSession()
        gdvBasketView = New DataView(gdtBasket)
        gdvBasketView.RowFilter = "ProductKey='" & sProductKey & "'"
        If gdvBasketView.Count > 0 Then
            gdvBasketView.Delete(0)
            Session(gsBasketCountName) = Session(gsBasketCountName) - 1
            SetBasketCount(Session(gsBasketCountName))
        End If
        gdvBasketView.RowFilter = ""
        Call SaveBasketToSession()
    End Sub
       
    Protected Sub GetConsignorAddress()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetCnorDetails", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int, 4)
        oParam.Value = CLng(Session("UserKey"))
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("CnorCompany")) Then
                psCnorCompany = oDataReader("CnorCompany")
            End If
            If Not IsDBNull(oDataReader("CnorAddr1")) Then
                psCnorAddr1 = oDataReader("CnorAddr1")
            End If
            If Not IsDBNull(oDataReader("CnorAddr2")) Then
                psCnorAddr2 = oDataReader("CnorAddr2")
            End If
            If Not IsDBNull(oDataReader("CnorAddr3")) Then
                psCnorAddr3 = oDataReader("CnorAddr3")
            End If
            If Not IsDBNull(oDataReader("CnorTown")) Then
                psCnorTown = oDataReader("CnorTown")
            End If
            If Not IsDBNull(oDataReader("CnorState")) Then
                psCnorState = oDataReader("CnorState")
            End If
            If Not IsDBNull(oDataReader("CnorPostCode")) Then
                psCnorPostCode = oDataReader("CnorPostCode")
            End If
            If Not IsDBNull(oDataReader("CnorCountryName")) Then
                psCnorCountryName = oDataReader("CnorCountryName")
            End If
            If Not IsDBNull(oDataReader("CnorCountryKey")) Then
                psCnorCountryKey = oDataReader("CnorCountryKey")
            End If
            If Not IsDBNull(oDataReader("CnorCtcName")) Then
                psCnorCtcName = oDataReader("CnorCtcName")
            End If
            If Not IsDBNull(oDataReader("CnorCtcTel")) Then
                psCnorCtcTel = oDataReader("CnorCtcTel")
            End If
            If Not IsDBNull(oDataReader("CnorCtcEmail")) Then
                psCnorCtcEmail = oDataReader("CnorCtcEmail")
            End If
            oDataReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetCountries()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Country_GetCountries", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Try
            oConn.Open()
            ddlCneeCountry.DataSource = oCmd.ExecuteReader()
            ddlCneeCountry.DataTextField = "CountryName"
            ddlCneeCountry.DataValueField = "CountryKey"
            ddlCneeCountry.DataBind()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function ValidBasket() As Boolean
        If IsShowingRichView() Then

            pdblBasketTotalValue = 0.0
            plBasketWeightGrams = 0

            Dim dgi As DataGridItem
            Dim txtOrderQuantity As TextBox
            Dim hidQtyAvailable As HiddenField
            Dim hidCalendarManaged As HiddenField
            Dim hidCustomLetter As HiddenField
            Dim lblQtyAvailable As Label
            Dim i As Integer
            Dim bValid As Boolean = True
            lblBasketMessage1.Text = String.Empty
            lblBasketMessage2.Text = String.Empty
            For Each dgi In gvBasket.Items
                txtOrderQuantity = CType(dgi.Cells(1).FindControl("txtOrderQuantity"), TextBox)
                hidQtyAvailable = CType(dgi.Cells(1).FindControl("hidQtyAvailable"), HiddenField)
                hidCalendarManaged = CType(dgi.Cells(1).FindControl("hidCalendarManaged"), HiddenField)
                hidCustomLetter = CType(dgi.Cells(1).FindControl("hidCustomLetter"), HiddenField)
                lblQtyAvailable = CType(dgi.Cells(1).FindControl("lblQtyAvailable"), Label)
                Dim sQtyToPick As String = txtOrderQuantity.Text.ToString
                Dim sQtyAvailable As String = hidQtyAvailable.Value.ToString
                Dim bIsCalendarManaged As Boolean = CBool(hidCalendarManaged.Value)
                Dim bIsCustomLetter As Boolean = CBool(hidCustomLetter.Value)
                i = i + 1
                If IsNumeric(sQtyToPick) AndAlso IsInteger(sQtyToPick) AndAlso IsNumeric(sQtyAvailable) Then
                    If Not bIsCustomLetter Then
                        If (CLng(sQtyToPick) > CLng(sQtyAvailable)) Then
                            bValid = False
                            If CInt(sQtyAvailable) = 0 Then
                                lblBasketMessage1.Text = "Item " & i & " is not available - please remove it from your basket before proceeding"
                            Else
                                lblBasketMessage1.Text = "You cannot order more than " & sQtyAvailable & " unit" & IIf(CInt(sQtyAvailable) = 1, "", "s") & " of item " & i & ". Please adjust your order quantity."
                            End If
                            lblBasketMessage2.Text = lblBasketMessage1.Text
                        ElseIf CLng(sQtyToPick) = 0 Then
                            bValid = False
                            lblBasketMessage1.Text = "Item " & i & " has an Order Quantity of zero - please adjust the quantity or remove the item before proceeding"
                            lblBasketMessage2.Text = lblBasketMessage1.Text
                        ElseIf CLng(sQtyToPick) < 0 Then
                            bValid = False
                            lblBasketMessage1.Text = "Item " & i & " has a negative Order Quantity - please adjust the quantity or remove the item before proceeding"
                            lblBasketMessage2.Text = lblBasketMessage1.Text
                        Else
                            Call GetBasketFromSession()
                            gdvBasketView = New DataView(gdtBasket)
                            gdvBasketView.RowFilter = "ProductKey='" & dgi.Cells(0).Text & "'"
                            If gdvBasketView.Count = 1 Then
                                With gdvBasketView(0)
                                    .Item("QtyToPick") = sQtyToPick

                                    pdblBasketTotalValue = pdblBasketTotalValue + (CDbl(.Item("UnitValue")) * CLng(sQtyToPick))
                                    lblPerCustomerConfiguration4Confirmation2BasketValue.Text = Format(pdblBasketTotalValue, "##,##0.00")  ' HYSTER CC
                                    plBasketWeightGrams = plBasketWeightGrams + (CLng(.Item("UnitWeightGrams")) * CLng(sQtyToPick))
                                    If ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedIndex > 0 Then
                                        Call GetShippingCosts()
                                    End If
                                End With
                            End If
                            gdvBasketView.RowFilter = ""
                            Call SaveBasketToSession()
                        End If
                    End If
                Else
                    bValid = False
                    lblBasketMessage1.Text = "Please ensure all order quantities have numeric values"
                    lblBasketMessage2.Text = lblBasketMessage1.Text
                End If
            Next dgi
            ValidBasket = bValid
        Else
            ValidBasket = ValidateClassicBasket()
        End If
    End Function
    
    Protected Function ValidateClassicBasket() As Boolean
        Dim dgi As DataGridItem
        Dim tbPickQty As TextBox
        Dim i As Integer
        Dim bValid As Boolean
        lblError.Text = ""
        bValid = True
        pdblBasketTotalValue = 0.0
        plBasketWeightGrams = 0
        For Each dgi In dgrdBasket.Items

            Call GetBasketFromSession()
            gdvBasketView = New DataView(gdtBasket)
            gdvBasketView.RowFilter = "ProductKey='" & dgi.Cells(0).Text & "'"
            Dim bCalendarManaged As Boolean
            Dim bCustomLetter As Boolean
            Dim nOnDemand As Integer
            Dim nOnDemandPriceList As Integer
            If gdvBasketView.Count > 0 Then
                bCalendarManaged = gdvBasketView(0).Item("CalendarManaged")
                bCustomLetter = gdvBasketView(0).Item("CustomLetter")
                nOnDemand = gdvBasketView(0).Item("OnDemand")
                nOnDemandPriceList = gdvBasketView(0).Item("OnDemandPriceList")
            End If
            
            tbPickQty = CType(dgi.Cells(10).FindControl("txtPickQuantity"), TextBox)
            Dim sQtyAvailable As String = dgi.Cells(8).Text
            Dim sQtyToPick As String = tbPickQty.Text.ToString
            Dim dr As DataRow
            i = i + 1
            If IsNumeric(sQtyToPick) And IsNumeric(sQtyAvailable) Then
                If (CLng(sQtyToPick) > CLng(sQtyAvailable)) And Not (bCalendarManaged Or bCustomLetter) Then
                    bValid = False
                    lblError.Text = "Row " & i & " has a pick quantity that exceeds the available quantity"
                ElseIf CLng(sQtyToPick) = 0 And Not (bCalendarManaged Or bCustomLetter) Then
                    bValid = False
                    lblError.Text = "Row " & i & " has a pick quantity of zero - please remove item before proceeding"
                ElseIf CLng(sQtyToPick) < 0 Then
                    bValid = False
                    lblError.Text = "Row " & i & " has a negative pick quantity - please remove item before proceeding"
                Else
                    If gdvBasketView.Count > 0 Then
                        gdvBasketView.Delete(0)
                    End If
                    gdvBasketView.RowFilter = ""
                    dr = gdtBasket.NewRow()
                    dr("ProductKey") = dgi.Cells(0).Text
                    dr("ProductCode") = dgi.Cells(2).Text
                    dr("ProductDate") = dgi.Cells(3).Text
                    dr("Description") = dgi.Cells(4).Text
                    dr("LanguageID") = dgi.Cells(5).Text
                    dr("UnitValue") = CDbl(dgi.Cells(6).Text.Substring(1))
                    dr("QtyAvailable") = CLng(sQtyAvailable)
                    dr("UnitWeightGrams") = dgi.Cells(9).Text
                    dr("QtyToPick") = CLng(sQtyToPick)

                    dr("BoxQty") = ""
                    dr("UnitValue2") = 0
                    dr("PDFFileName") = ""
                    dr("OriginalImage") = ""
                    dr("ThumbNailImage") = ""
                    dr("Notes") = ""
                    dr("CalendarManaged") = bCalendarManaged
                    dr("OnDemand") = nOnDemand
                    dr("OnDemandPriceList") = nOnDemandPriceList
                    dr("CustomLetter") = bCustomLetter
                    gdtBasket.Rows.Add(dr)
                    Call SaveBasketToSession()
                    pdblBasketTotalValue = pdblBasketTotalValue + (CDbl(dr("UnitValue")) * CLng(sQtyToPick))
                    lblPerCustomerConfiguration4Confirmation2BasketValue.Text = Format(pdblBasketTotalValue, "##,##0.00")  ' HYSTER CC
                    plBasketWeightGrams = plBasketWeightGrams + (CLng(dr("UnitWeightGrams")) * CLng(sQtyToPick))
                    If ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedIndex > 0 Then
                        Call GetShippingCosts()
                    End If
                End If
            Else
                bValid = False
                lblError.Text = "please ensure that all pick quantities have numeric values"
            End If
        Next dgi
        If bValid Then
            ValidateClassicBasket = True
        Else
            ValidateClassicBasket = False
        End If
    End Function
    
    Function IsInteger(ByVal sValue As String) As Boolean
        sValue = sValue.Trim
        IsInteger = True
        For i As Integer = 0 To sValue.Length - 1
            If sValue.Chars(i) < "0" Or sValue.Chars(i) > "9" Then
                IsInteger = False
                Exit Function
            End If
        Next
    End Function
    
    Protected Sub SaveRetrievedAddress()
        psRetrievedAddress = CompressCneeAddress()
    End Sub

    Protected Function CompressCneeAddress() As String
        Dim sbAddr As New StringBuilder
        sbAddr.Append(txtCneeName.Text)
        sbAddr.Append(txtCneeAddr1.Text)
        sbAddr.Append(txtCneeAddr2.Text)
        sbAddr.Append(txtCneeAddr3.Text)
        sbAddr.Append(txtCneeCity.Text)
        sbAddr.Append(txtCneeState.Text)
        sbAddr.Append(txtCneePostCode.Text)
        sbAddr.Append(ddlCneeCountry.SelectedItem.Text)
        sbAddr.Append(txtCneeCtcName.Text)
        sbAddr.Append(txtCneeTel.Text)
        sbAddr.Append(txtCneeEmail.Text)
        CompressCneeAddress = sbAddr.ToString
    End Function
    
    Protected Sub GetConsigneeAddress()
        If plCneeAddressKey > 0 Then
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_GlobalAddress_GetFromKey", oConn)
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
                    txtCneeName.Text = oDataReader("Company").ToString.Trim
                End If
                If IsDBNull(oDataReader("Addr1")) Then
                    txtCneeAddr1.Text = ""
                Else
                    txtCneeAddr1.Text = oDataReader("Addr1").ToString.Trim
                End If
                If IsDBNull(oDataReader("Addr2")) Then
                    txtCneeAddr2.Text = ""
                Else
                    txtCneeAddr2.Text = oDataReader("Addr2").ToString.Trim
                End If
                If IsDBNull(oDataReader("Addr3")) Then
                    txtCneeAddr3.Text = ""
                Else
                    txtCneeAddr3.Text = oDataReader("Addr3").ToString.Trim
                End If
                If IsDBNull(oDataReader("Town")) Then
                    txtCneeCity.Text = ""
                Else
                    txtCneeCity.Text = oDataReader("Town").ToString.Trim
                End If
                If IsDBNull(oDataReader("State")) Then
                    txtCneeState.Text = ""
                Else
                    txtCneeState.Text = oDataReader("State").ToString.Trim
                End If
                If IsDBNull(oDataReader("PostCode")) Then
                    txtCneePostCode.Text = ""
                Else
                    txtCneePostCode.Text = oDataReader("PostCode").ToString.Trim
                End If
                If Not IsDBNull(oDataReader("CountryKey")) Then
                    Call SetCountryDropdown(oDataReader("CountryKey"))
                    Call SetCountry(oDataReader("CountryKey"), oDataReader("State").ToString.Trim & String.Empty)
                End If
                'If IsDBNull(oDataReader("CountryName")) Then
                ' ddlCneeCountry.SelectedItem.Text = ""
                ' Else
                ' ddlCneeCountry.SelectedItem.Text = oDataReader("CountryName")
                ' End If
                If IsDBNull(oDataReader("AttnOf")) Then
                    txtCneeCtcName.Text = ""
                Else
                    txtCneeCtcName.Text = oDataReader("AttnOf").ToString.Trim
                End If
                If IsDBNull(oDataReader("Telephone")) Then
                    txtCneeTel.Text = ""
                Else
                    txtCneeTel.Text = oDataReader("Telephone").ToString.Trim
                End If
                If IsDBNull(oDataReader("Email")) Then
                    txtCneeEmail.Text = ""
                Else
                    txtCneeEmail.Text = oDataReader("Email").ToString.Trim
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            Call SaveRetrievedAddress()
        End If
    End Sub
    
    Protected Sub SetCountryDropdown(ByVal sCountryKey As String)
        If IsNumeric(sCountryKey) Then
            Dim nCountryKey As Integer = CInt(sCountryKey)
            For i As Integer = 0 To ddlCneeCountry.Items.Count - 1
                If ddlCneeCountry.Items(i).Value = nCountryKey Then
                    ddlCneeCountry.SelectedIndex = i
                    Call SetCountry(ddlCneeCountry.SelectedValue, "")
                    Exit For
                End If
            Next
        End If
    End Sub
    
    Protected Sub AddNewAddress()
        Dim bError As Boolean = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_Add2", oConn)
    
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim p1 As Integer = Session("CustomerKey")

        Dim paramCompany As SqlParameter = New SqlParameter("@Company", SqlDbType.NVarChar, 50)
        paramCompany.Value = txtCneeName.Text
        oCmd.Parameters.Add(paramCompany)
        
        Dim p3 As String = txtCneeName.Text
        
        Dim paramAddr1 As SqlParameter = New SqlParameter("@Addr1", SqlDbType.NVarChar, 50)
        paramAddr1.Value = txtCneeAddr1.Text
        oCmd.Parameters.Add(paramAddr1)
        
        Dim p4 As String = txtCneeAddr1.Text

        Dim paramparamAddr2 As SqlParameter = New SqlParameter("@Addr2", SqlDbType.NVarChar, 50)
        paramparamAddr2.Value = txtCneeAddr2.Text
        oCmd.Parameters.Add(paramparamAddr2)
        
        Dim p5 As String = txtCneeAddr2.Text
        
        Dim paramparamAddr3 As SqlParameter = New SqlParameter("@Addr3", SqlDbType.NVarChar, 50)
        paramparamAddr3.Value = txtCneeAddr3.Text
        oCmd.Parameters.Add(paramparamAddr3)
        
        Dim p6 As String = txtCneeAddr3.Text
        
        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = txtCneeCity.Text
        oCmd.Parameters.Add(paramTown)
        
        Dim p7 As String = txtCneeCity.Text
        
        Dim paramState As SqlParameter = New SqlParameter("@State", SqlDbType.NVarChar, 50)
        paramState.Value = txtCneeState.Text
        oCmd.Parameters.Add(paramState)
        
        Dim p8 As String = txtCneeState.Text
        
        Dim paramPostCode As SqlParameter = New SqlParameter("@PostCode", SqlDbType.NVarChar, 50)
        paramPostCode.Value = txtCneePostCode.Text
        oCmd.Parameters.Add(paramPostCode)
        
        Dim p9 As String = txtCneePostCode.Text
        
        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int, 4)
        paramCountryKey.Value = CLng(ddlCneeCountry.SelectedItem.Value)
        oCmd.Parameters.Add(paramCountryKey)
        
        Dim p10 As String = ddlCneeCountry.SelectedItem.Value
        
        Dim paramAttnOf As SqlParameter = New SqlParameter("@AttnOf", SqlDbType.NVarChar, 50)
        paramAttnOf.Value = txtCneeCtcName.Text
        oCmd.Parameters.Add(paramAttnOf)
        
        Dim p13 As String = txtCneeCtcName.Text

        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 50)
        paramTelephone.Value = txtCneeTel.Text
        oCmd.Parameters.Add(paramTelephone)
        
        Dim p14 As String = txtCneeTel.Text
        
        Dim paramEmail As SqlParameter = New SqlParameter("@Email", SqlDbType.NVarChar, 50)
        paramEmail.Value = txtCneeEmail.Text
        oCmd.Parameters.Add(paramEmail)
        
        Dim p16 As String = txtCneeEmail.Text

        Dim paramLastUpdatedByKey As SqlParameter = New SqlParameter("@LastUpdatedByKey", SqlDbType.Int, 4)
        paramLastUpdatedByKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramLastUpdatedByKey)
        
        Dim p17 As String = Session("UserKey")
        
        Dim paramAddressKey As SqlParameter = New SqlParameter("@AddressKey", SqlDbType.Int, 4)
        paramAddressKey.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramAddressKey)
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            If Not paramAddressKey.Value Is Nothing Then
                plCneeAddressKey = paramAddressKey.Value
            Else
                plCneeAddressKey = 0
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
            plCneeAddressKey = 0
            oConn.Close()
            bError = True
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub AddToPersonalAddressBook()
        If plCneeAddressKey > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToPersonal", oConn)
            'Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
            paramUserKey.Value = Session("UserKey")
            oCmd.Parameters.Add(paramUserKey)
            Dim paramGABKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramGABKey.Value = plCneeAddressKey
            oCmd.Parameters.Add(paramGABKey)
            Try
                oConn.Open()
                oCmd.Connection = oConn
                'oCmd.Transaction = oTrans
                oCmd.ExecuteNonQuery()
                'oTrans.Commit()
            Catch ex As SqlException
                'oTrans.Rollback("AddRecord")
                lblError.Text = ex.ToString
                plCneeAddressKey = 0
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Sub AddToSharedAddressBook()
        If plCneeAddressKey > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Address_AddToGlobal", oConn)
            'Dim oTrans As SqlTransaction
            oCmd.CommandType = CommandType.StoredProcedure
            Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            paramCustomerKey.Value = Session("CustomerKey")
            oCmd.Parameters.Add(paramCustomerKey)
            Dim paramAddressKey As SqlParameter = New SqlParameter("@GlobalAddressKey", SqlDbType.Int, 4)
            paramAddressKey.Value = plCneeAddressKey
            oCmd.Parameters.Add(paramAddressKey)
            Try
                oConn.Open()
                oCmd.Connection = oConn
                oCmd.ExecuteNonQuery()
            Catch ex As SqlException
                plCneeAddressKey = 0
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Function sBuildConsignorAddress() As String
        Dim sConsignor As String
        Dim CR As String = "<br/>"
        sConsignor = psCnorCompany & CR
        sConsignor &= psCnorAddr1 & CR
        If psCnorAddr2 <> "" Then
            If psCnorAddr3 <> "" Then
                sConsignor &= psCnorAddr2 & " " & psCnorAddr3 & CR
            Else
                sConsignor &= psCnorAddr2 & CR
            End If
        End If
        If psCnorState <> "" Then
            If psCnorPostCode <> "" Then
                sConsignor &= psCnorTown & " " & psCnorState & " " & psCnorPostCode & CR
            Else
                sConsignor &= psCnorTown & " " & psCnorState & CR
            End If
        Else
            If psCnorPostCode <> "" Then
                sConsignor &= psCnorTown & " " & psCnorPostCode & CR
            Else
                sConsignor &= psCnorTown & CR
            End If
        End If
        sConsignor &= psCnorCountryName & CR
        If psCnorCtcTel <> "" Then
            sConsignor &= psCnorCtcName & " " & psCnorCtcTel & CR
        Else
            sConsignor &= psCnorCtcName & CR
        End If
        If psCnorCtcEmail <> "" Then
            sConsignor &= psCnorCtcEmail
        End If
        sBuildConsignorAddress = sConsignor
    End Function
    
    Protected Sub SummariseBooking()
        Call GetBasketFromSession()
        If Not IsNothing(gdtBasket) Then
            Dim sConsignor, sConsignee As String
            Dim CR As String = "<br/>"
            lblCheckOutMessage.Text = ""
            
            sConsignor = psCnorCompany & CR
            sConsignor &= psCnorAddr1 & CR
            If psCnorAddr2 <> "" Then
                If psCnorAddr3 <> "" Then
                    sConsignor &= psCnorAddr2 & " " & psCnorAddr3 & CR
                Else
                    sConsignor &= psCnorAddr2 & CR
                End If
            End If
            If psCnorState <> "" Then
                If psCnorPostCode <> "" Then
                    sConsignor &= psCnorTown & " " & psCnorState & " " & psCnorPostCode & CR
                Else
                    sConsignor &= psCnorTown & " " & psCnorState & CR
                End If
            Else
                If psCnorPostCode <> "" Then
                    sConsignor &= psCnorTown & " " & psCnorPostCode & CR
                Else
                    sConsignor &= psCnorTown & CR
                End If
            End If
            sConsignor &= psCnorCountryName & CR
            If psCnorCtcTel <> "" Then
                sConsignor &= psCnorCtcName & " " & psCnorCtcTel & CR
            Else
                sConsignor &= psCnorCtcName & CR
            End If
            If psCnorCtcEmail <> "" Then
                sConsignor &= psCnorCtcEmail
            End If
            
            sConsignee = Session("SB_CneeCompany") & CR
            sConsignee &= Session("SB_CneeAddr1") & CR
            If Session("SB_CneeAddr2") <> "" Then
                If Session("SB_CneeAddr3") <> "" Then
                    sConsignee &= Session("SB_CneeAddr2") & " " & Session("SB_CneeAddr3") & CR
                Else
                    sConsignee &= Session("SB_CneeAddr2") & CR
                End If
            End If
            If Session("SB_CneeState") <> "" Then
                If Session("SB_CneePostCode") <> "" Then
                    sConsignee &= Session("SB_CneeTown") & " " & Session("SB_CneeState") & " " & Session("SB_CneePostCode") & CR
                Else
                    sConsignee &= Session("SB_CneeTown") & " " & Session("SB_CneeState") & CR
                End If
            Else
                If Session("SB_CneePostCode") <> "" Then
                    sConsignee &= Session("SB_CneeTown") & " " & Session("SB_CneePostCode") & CR
                Else
                    sConsignee &= Session("SB_CneeTown") & CR
                End If
            End If
            sConsignee &= Session("SB_CneeCountryName") & CR
            If Session("SB_CneeCtcTel") <> "" Then
                sConsignee &= Session("SB_CneeCtcName") & " " & Session("SB_CneeCtcTel") & CR
            Else
                sConsignee &= Session("SB_CneeCtcName") & CR
            End If
            If Session("SB_CneeCtcEmail") <> "" Then
                sConsignee &= Session("SB_CneeCtcEmail")
            End If
            
            lblConsignor.Text = sConsignor
            lblConsignee.Text = sConsignee
        
            Call GetBasketFromSession()
            gdvBasketView = New DataView(gdtBasket)
            gdvBasketView.Sort = "ProductCode"
            gvConfirmationBasket.DataSource = gdvBasketView
            gvConfirmationBasket.DataBind()
            gvConfirmationBasket.Visible = True
            If IsCAB() Then
                Dim dblMaterialsSum As Double = 0.0
                For Each dr As DataRow In gdtBasket.Rows
                    dblMaterialsSum += dr("UnitValue") * dr("QtyToPick")
                Next
                lblMaterialsSummary.Text = "PRODUCT COST TOTAL:&nbsp;&nbsp;&nbsp;  £" & Format(dblMaterialsSum, "##,##0.00")
            End If
            
            lblCustRef1.Text = Session("SB_BookingRef1")
            lblCustRef2.Text = Session("SB_BookingRef2")
            lblCustRef3.Text = Session("SB_BookingRef3")
            lblCustRef4.Text = Session("SB_BookingRef4")
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_1_BLACKROCK Then
                lblPerCustomerConfiguration1ConfirmationCostCentre.Text = Session("SB_BookingRef3")
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3 Or _
              plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Or _
                plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_16_VSOE Then
                lblPerCustomerConfiguration2ConfirmationBookingRef.Text = Session("SB_BookingRef3")
                lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefA.Text = Session("SB_BookingRef4")
                lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefB.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefC.Text = Session("SB_BookingRef2")
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_3_LEGACY_SINGLE_MANDATORY_UNPROMPTED_COST_CENTRE Then
                lblPerCustomerConfiguration3ConfirmationCostCentre.Text = Session("SB_BookingRef3")
                lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefA.Text = Session("SB_BookingRef4")
                lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefB.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefC.Text = Session("SB_BookingRef2")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_5_KODAK Then
                lblPerCustomerConfiguration5ConfirmationCostCentre.Text = Session("SB_BookingRef3")
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_6_CIMA Then
                lblPerCustomerConfiguration6ConfirmationDepartment.Text = Session("SB_BookingRef1")  ' added CN
                lblPerCustomerConfiguration6ConfirmationCostCentre.Text = Session("SB_BookingRef2")  ' added CN
                lblPerCustomerConfiguration6ConfirmationReference.Text = Session("SB_BookingRef3")  ' added CN
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Then
                lblPerCustomerConfiguration7ConfirmationServiceLevel.Text = psServiceLevel
            End If
            
            If IsVSAL() Then
                lblPerCustomerConfiguration18ConfirmationServiceLevel.Text = psServiceLevel
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
                lblPerCustomerConfiguration8ConfirmationBookingRef.Text = Session("SB_BookingRef")
                lblPerCustomerConfiguration8ConfirmationPCID.Text = Session("SB_PCID")
                lblPerCustomerConfiguration8ConfirmationRating.Text = Session("SB_Rating")
                lblPerCustomerConfiguration8ConfirmationRO.Text = Session("SB_RO")
                lblPerCustomerConfiguration8ConfirmationMDSOrderRef.Text = Session("SB_MDSOrderRef")
            End If
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_12_ATKINS Then
                lblPerCustomerConfiguration12ConfirmationCostCentre.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefA.Text = Session("SB_BookingRef2")
                lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefB.Text = Session("SB_BookingRef3")
                lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefC.Text = Session("SB_BookingRef4")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_19_INSIGHT Then
                lblPerCustomerConfiguration19ConfirmationCostCentre.Text = Session("SB_BookingRef1")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_20_DAT Then
                lblPerCustomerConfiguration20ConfirmationCostCentre.Text = Session("SB_BookingRef1")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_21_UNICRD Then
                lblPerCustomerConfiguration21ConfirmationCostCentre.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration21ConfirmationCustRef.Text = Session("SB_BookingRef2")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_22_RIOTINTO Then
                lblPerCustomerConfiguration22ConfirmationCostCentre.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration22ConfirmationRequestedBy.Text = Session("SB_BookingRef2")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_23_ARTHRITIS Then
                lblPerCustomerConfiguration23ConfirmationCostCentre.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration23ConfirmationCategory.Text = Session("SB_BookingRef2")
                lblPerCustomerConfiguration23ConfirmationPONumber.Text = Session("SB_BookingRef3")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_24_PROMOVERITAS Then
                lblPerCustomerConfiguration24ConfirmationReference.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration24ConfirmationRequestedBy.Text = Session("SB_BookingRef2")
                lblPerCustomerConfiguration24ConfirmationRecipientName.Text = Session("SB_BookingRef3")
                lblPerCustomerConfiguration24ConfirmationProducts.Text = Session("SB_BookingRef4")
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_29_IRWINMITCHELL Then
                lblPerCustomerConfiguration29ConfirmationBudgetCode.Text = Session("SB_BookingRef1")
                lblPerCustomerConfiguration29ConfirmationDepartment.Text = Session("SB_BookingRef2")
                lblPerCustomerConfiguration29ConfirmationEvent.Text = Session("SB_BookingRef3")
                lblPerCustomerConfiguration29ConfirmationBDName.Text = Session("SB_BookingRef4")
            End If

            If IsJupiter() Then
                lblPerCustomerConfirmation30BudgetCode.Text = ddlPerCustomerConfiguration30JupiterBudgetCode.SelectedItem.Text
            End If
            
            lblSpecialInstructions.Text = Session("SB_SpecialInstructions")
            lblShippingInfo.Text = Session("SB_ShippingNote")
        Else
            lblCheckOutMessage.Text = "please choose one or more products before proceeding to checkout"
        End If
    End Sub
    
    Protected Sub CaptureBookingInstructions()
        Session("SB_CneeCompany") = txtCneeName.Text
        Session("SB_CneeAddr1") = txtCneeAddr1.Text
        Session("SB_CneeAddr2") = txtCneeAddr2.Text
        Session("SB_CneeAddr3") = txtCneeAddr3.Text
        Session("SB_CneeTown") = txtCneeCity.Text
        Session("SB_CneeState") = txtCneeState.Text
        Session("SB_CneePostCode") = txtCneePostCode.Text
        Session("SB_CneeCountryKey") = ddlCneeCountry.SelectedItem.Value
        Session("SB_CneeCountryName") = ddlCneeCountry.SelectedItem.Text
        Session("SB_CneeCtcName") = txtCneeCtcName.Text
        Session("SB_CneeCtcTel") = txtCneeTel.Text
        Session("SB_CneeCtcEmail") = txtCneeEmail.Text
        Session("SB_BookingRef1") = txtCustRef1.Text
        Session("SB_BookingRef2") = txtCustRef2.Text
        Session("SB_BookingRef3") = txtCustRef3.Text
        Session("SB_BookingRef4") = txtCustRef4.Text
        Session("SB_SpecialInstructions") = txtSpecialInstructions.Text
        Session("SB_ShippingNote") = txtShippingInfo.Text
        
        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_1_BLACKROCK Then
            Session("SB_BookingRef3") = tbPerCustomerConfiguration1CostCentre.Text
            Session("SB_BookingRef4") = String.Empty
            Session("SB_BookingRef1") = String.Empty
            Session("SB_BookingRef2") = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3 Or _
          plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Or _
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_16_VSOE Then
            Session("SB_BookingRef3") = tbPerCustomerConfiguration2BookingRef.Text
            Session("SB_BookingRef4") = tbPerCustomerConfiguration2AdditionalRefA.Text
            Session("SB_BookingRef1") = tbPerCustomerConfiguration2AdditionalRefB.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration2AdditionalRefC.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_3_LEGACY_SINGLE_MANDATORY_UNPROMPTED_COST_CENTRE Then
            Session("SB_BookingRef3") = tbPerCustomerConfiguration3CostCentre.Text
            Session("SB_BookingRef4") = tbPerCustomerConfiguration3AdditionalRefA.Text
            Session("SB_BookingRef1") = tbPerCustomerConfiguration3AdditionalRefB.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration3AdditionalRefC.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_5_KODAK Then
            Session("SB_BookingRef3") = tbPerCustomerConfiguration5CostCentre.Text
            Session("SB_BookingRef4") = String.Empty
            Session("SB_BookingRef1") = String.Empty
            Session("SB_BookingRef2") = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_6_CIMA Then
            Session("SB_BookingRef1") = ddlPerCustomerConfiguration6Department.SelectedItem.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration6CostCentre.Text
            Session("SB_BookingRef3") = tbPerCustomerConfiguration6Reference.Text
            Session("SB_BookingRef4") = Session("UserName")
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Then
            psServiceLevel = ddlPerCustomerConfiguration7ServiceLevel.SelectedValue
        End If
        
        If IsVSAL() Then
            psServiceLevel = ddlPerCustomerConfiguration18ServiceLevel.SelectedValue
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
            Session("SB_BookingRef") = tbPerCustomerConfiguration8BookingRef.Text
            Session("SB_RO") = ddlPerCustomerConfiguration8RO.Text
            Session("SB_Rating") = ddlPerCustomerConfiguration8Rating.Text
            Session("SB_PCID") = tbPerCustomerConfiguration8PCID.Text
            Session("SB_MDSOrderRef") = tbPerCustomerConfiguration8MDSOrderRef.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_12_ATKINS Then
            Session("SB_BookingRef1") = tbPerCustomerConfiguration12CostCentre.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration12AdditionalRefA.Text
            Session("SB_BookingRef3") = tbPerCustomerConfiguration12AdditionalRefB.Text
            Session("SB_BookingRef4") = tbPerCustomerConfiguration12AdditionalRefC.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_19_INSIGHT Then
            Session("SB_BookingRef1") = tbPerCustomerConfiguration19CostCentre.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_20_DAT Then
            Session("SB_BookingRef1") = ddlPerCustomerConfiguration20CostCentre.SelectedItem.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_21_UNICRD Then
            Session("SB_BookingRef1") = ddlPerCustomerConfiguration21CostCentre.SelectedItem.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration21CustRef.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_22_RIOTINTO Then
            Session("SB_BookingRef1") = tbPerCustomerConfiguration22CostCentre.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration22RequestedBy.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_23_ARTHRITIS Then
            Session("SB_BookingRef1") = ddlPerCustomerConfiguration23CostCentre.SelectedValue
            Session("SB_BookingRef2") = ddlPerCustomerConfiguration23Category.SelectedValue
            Session("SB_BookingRef3") = tbPerCustomerConfiguration23PONumber.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_24_PROMOVERITAS Then
            Session("SB_BookingRef1") = tbPerCustomerConfiguration24Reference.Text
            Session("SB_BookingRef2") = tbPerCustomerConfiguration24RequestedBy.Text
            Session("SB_BookingRef3") = tbPerCustomerConfiguration24RecipientName.Text
            Session("SB_BookingRef4") = tbPerCustomerConfiguration24Products.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_29_IRWINMITCHELL Then
            Session("SB_BookingRef1") = ddlPerCustomerConfiguration29IrwinMitchellBudgetCode.SelectedItem.Text
            Session("SB_BookingRef2") = ddlPerCustomerConfiguration29IrwinMitchellDepartment.SelectedItem.Text
            Session("SB_BookingRef3") = tbPerCustomerConfiguration29Event.Text
            Session("SB_BookingRef4") = tbPerCustomerConfiguration29BDName.Text
        End If
    End Sub
    
    Protected Sub PaintSessionVariables()
        txtCneeName.Text = Session("SB_CneeCompany")
        txtCneeAddr1.Text = Session("SB_CneeAddr1")
        txtCneeAddr2.Text = Session("SB_CneeAddr2")
        txtCneeAddr3.Text = Session("SB_CneeAddr3")
        txtCneeCity.Text = Session("SB_CneeTown")
        txtCneeState.Text = Session("SB_CneeState")
        txtCneePostCode.Text = Session("SB_CneePostCode")
        Call SetCountryDropdown(Session("SB_CneeCountryKey"))
        If IsNumeric(Session("SB_CneeCountryKey")) Then
            Call SetCountry(Session("SB_CneeCountryKey"), Session("SB_CneeState"))
        Else
            Call SetCountry(0, Session("SB_CneeState"))
        End If
        'ddlCneeCountry.SelectedItem.Value = Session("SB_CneeCountryKey")
        'ddlCneeCountry.SelectedItem.Text = Session("SB_CneeCountryName")
        txtCneeCtcName.Text = Session("SB_CneeCtcName")
        txtCneeTel.Text = Session("SB_CneeCtcTel")
        txtCneeEmail.Text = Session("SB_CneeCtcEmail")
        txtCustRef1.Text = Session("SB_BookingRef1")
        txtCustRef2.Text = Session("SB_BookingRef2")
        txtCustRef3.Text = Session("SB_BookingRef3")
        txtCustRef4.Text = Session("SB_BookingRef4")
        txtSpecialInstructions.Text = Session("SB_SpecialInstructions")
        txtShippingInfo.Text = Session("SB_ShippingNote")

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3 Or _
          plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Or _
            plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_16_VSOE Then
            tbPerCustomerConfiguration2BookingRef.Text = Session("SB_BookingRef3")
            tbPerCustomerConfiguration2AdditionalRefA.Text = Session("SB_BookingRef4")
            tbPerCustomerConfiguration2AdditionalRefB.Text = Session("SB_BookingRef1")
            tbPerCustomerConfiguration2AdditionalRefC.Text = Session("SB_BookingRef2")
        End If
    End Sub
    
    Protected Sub ResetFields()
        txtCneeName.Text = String.Empty
        txtCneeAddr1.Text = String.Empty
        txtCneeAddr2.Text = String.Empty
        txtCneeAddr3.Text = String.Empty
        txtCneeCity.Text = String.Empty
        txtCneeState.Text = String.Empty
        txtCneePostCode.Text = String.Empty
        ddlCneeCountry.SelectedIndex = -1
        txtCneeCtcName.Text = String.Empty
        txtCneeTel.Text = String.Empty
        txtCneeEmail.Text = String.Empty
        txtCustRef1.Text = String.Empty
        txtCustRef2.Text = String.Empty
        txtCustRef3.Text = String.Empty
        txtCustRef4.Text = String.Empty
        txtSpecialInstructions.Text = String.Empty
        txtShippingInfo.Text = String.Empty

        tbPerCustomerConfiguration1CostCentre.Text = String.Empty
        ddlPerCustomerConfiguration1CostCentre.SelectedIndex = -1

        tbPerCustomerConfiguration2BookingRef.Text = String.Empty
        tbPerCustomerConfiguration2AdditionalRefA.Text = String.Empty
        tbPerCustomerConfiguration2AdditionalRefB.Text = String.Empty
        tbPerCustomerConfiguration2AdditionalRefC.Text = String.Empty

        tbPerCustomerConfiguration5CostCentre.Text = String.Empty
        ddlPerCustomerConfiguration5CostCentre.SelectedIndex = -1

        tbPerCustomerConfiguration6CostCentre.Text = String.Empty
        tbPerCustomerConfiguration6Reference.Text = String.Empty
        ddlPerCustomerConfiguration6Department.SelectedIndex = -1
        
        ddlPerCustomerConfiguration7ServiceLevel.SelectedIndex = -1
        
        tbPerCustomerConfiguration8BookingRef.Text = String.Empty
        tbPerCustomerConfiguration8PCID.Text = String.Empty
        ddlPerCustomerConfiguration8Rating.SelectedIndex = -1
        ddlPerCustomerConfiguration8RO.SelectedIndex = -1
        tbPerCustomerConfiguration8MDSOrderRef.Text = String.Empty

        tbPerCustomerConfiguration12CostCentre.Text = String.Empty
        tbPerCustomerConfiguration12AdditionalRefA.Text = String.Empty
        tbPerCustomerConfiguration12AdditionalRefB.Text = String.Empty
        tbPerCustomerConfiguration12AdditionalRefC.Text = String.Empty
        
        ddlPerCustomerConfiguration21CostCentre.SelectedIndex = -1
        tbPerCustomerConfiguration21CustRef.Text = String.Empty
        
        tbPerCustomerConfiguration22CostCentre.Text = String.Empty
        tbPerCustomerConfiguration22RequestedBy.Text = String.Empty

        trFinalCheckDefault.Visible = True
        trFinalCheckOrderAuthorisation.Visible = False
        tbAuthoriserMessageSingleAddressOrder.Text = String.Empty
        
        rblPerCustomerConfiguration0CheckoutNextDayDelivery00.Checked = True
        Call PerCustomerConfiguration0CheckoutNextDayDeliverySetCalendarVisibility()
        
        If IsAAT() Then
            Call InitUserCostCentreForAAT()
        End If
        
        If IsProquest() Then
            Call SetProquestMessage()
        End If
        
        ' SHOULD WE NOT RESET ALL POSSIBLE DROPDOWNS?
    End Sub
    
    Protected Sub ClearOrderSessionVariables()
        'Dim lstSessionVal As New System.Collections.Generic.List(Of String)
        'For Each s As String In Session.Keys
        ' If s.Substring(0, 3) = "SB_" Then lstSessionVal.Add(s)
        ' Next
        'For Each s As String In lstSessionVal
        ' Session.Remove(s)
        ' Next
        Session(gsBasketCountName) = Nothing
        Session("SB_BasketData") = Nothing
        Session("SB_BasketDataJupiter") = Nothing
        Session("SB_AuthorisationUsage") = Nothing

        Session("SB_BookingRef1") = String.Empty
        Session("SB_BookingRef2") = String.Empty
        Session("SB_BookingRef3") = String.Empty
        Session("SB_BookingRef4") = String.Empty
        Session("SB_SpecialInstructions") = String.Empty
        Session("SB_ShippingNote") = String.Empty

        Session("SB_CneeCompany") = String.Empty
        Session("SB_CneeAddr1") = String.Empty
        Session("SB_CneeAddr2") = String.Empty
        Session("SB_CneeAddr3") = String.Empty
        Session("SB_CneeTown") = String.Empty
        Session("SB_CneeState") = String.Empty
        Session("SB_CneePostCode") = String.Empty
        Session("SB_CneeCountryKey") = String.Empty
        Session("SB_CneeCtcName") = String.Empty
        Session("SB_CneeCtcTel") = String.Empty
        Session("SB_CneeCtcEmail") = String.Empty
        
        psServiceLevel = ""

        SetBasketCount("0")
        pnlAssociatedProducts.Visible = False
        If Not gdtMultiAddressBooking Is Nothing Then
            gdtMultiAddressBooking = Nothing
        End If
        Call ResetFields()
        pbAuthorisationRequired = False
        Call ShowSaveAddressLinks()
    End Sub

    Protected Sub btn_SelectAddress_click(ByVal s As Object, ByVal e As EventArgs)
        lblError.Text = ""
        Call GetBasketFromSession()  ' CN new
        If Not IsNothing(gdtBasket) Then
            Call ResetFields()
            ShowSearchAddressListPanel()
        Else
            lblError.Text = "Your basket is empty - add items to your basket before selecting the destination"
        End If
    End Sub
    
    Protected Sub btn_ConfirmAddress_click(ByVal s As Object, ByVal e As EventArgs)
        lblError.Text = ""
        Call GetBasketFromSession()  ' CN new
        If Not IsNothing(gdtBasket) Then
            ShowDeliveryAddressPanel()
        Else
            lblError.Text = "Your basket is empty - add items to your basket before completing your order"
        End If
    End Sub
    
    Protected Function IsValidMDSOrderRef() As Boolean
        IsValidMDSOrderRef = False
        Dim sMDSOrderRef As String = tbPerCustomerConfiguration8MDSOrderRef.Text.Trim
        If sMDSOrderRef.Length > 0 Then
            Dim oConn As New SqlConnection(gsConn)
            Dim oDataTable As New DataTable
            Dim sProc As String = "MDS_Order_CountOrdersWithMDSOrderRef"
            Dim oAdapter As New SqlDataAdapter(sProc, oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@MDSOrderRef", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@MDSOrderRef").Value = sMDSOrderRef
            oAdapter.Fill(oDataTable)
            IsValidMDSOrderRef = CInt(oDataTable.Rows(0).Item(0)) = 0
        End If
    End Function
    
    Protected Sub btn_CheckOut_click(ByVal s As Object, ByVal e As System.EventArgs)
        btnViewInvoice.Visible = False
        If ValidBasket() Then
            If pbCalendarManagement AndAlso (Not OneOnlyOfEachCMProductRequested()) Then
                WebMsgBox.Show("You cannot select more than one of a Calendar Managed product.\n\nEach physical Calendar Managed item has a unique product code. To select more than one item of the same type, select more than one product. For more information contact your Account Handler.")
                Exit Sub
            End If
            Call CheckOut()
            If IsWURS() Then
                Call CheckForWURSCriticalProducts()
            End If
        Else
            WebMsgBox.Show("Invalid basket.")
        End If
    End Sub
    
    Protected Function OneOnlyOfEachCMProductRequested() As Boolean
        OneOnlyOfEachCMProductRequested = True
        For Each dr As DataRow In gdtBasket.Rows
            If dr("CalendarManaged") = True AndAlso CInt(dr("QtyToPick")) > 1 Then
                OneOnlyOfEachCMProductRequested = False
                Exit For
            End If
        Next
    End Function
    
    Protected Function GetPricePerItem(ByVal nTariffId As Integer, ByVal nQuantityOrdered As Integer) As Double
        Dim sSQL As String
        sSQL = "SELECT * FROM OnDemandTariff WHERE TariffId = " & nTariffId & " ORDER BY Quantity"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim dblLastPrice As Double = -1
        For Each dr As DataRow In oDataTable.Rows
            dblLastPrice = dr("Price")
            If nQuantityOrdered <= dr("Quantity") Then
                Exit For
            End If
        Next
        GetPricePerItem = dblLastPrice
    End Function
    
    Protected Sub CheckForWURSCriticalProducts()
        pbBasketContainsWURSCriticalProducts = False
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT Misc2 FROM LogisticProduct WHERE ")
        Call GetBasketFromSession()
        For Each dr As DataRow In gdtBasket.Rows
            sbSQL.Append("LogisticProductKey = " & dr("ProductKey").ToString & " OR ")
        Next
        sbSQL.Remove(sbSQL.Length - 4, 4)
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read()
                If Not IsDBNull(oDataReader("Misc2")) Then
                    If oDataReader("Misc2").ToString.Trim.ToLower = "y" Then
                        pbBasketContainsWURSCriticalProducts = True
                        Exit While
                    End If
                End If
            End While
        Catch ex As Exception
            WebMsgBox.Show("Error in CheckForWURSCriticalProducts: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub CheckOut()
        Call GetBasketFromSession()  'CN new
        If gnMode = MODE_JUPITER_POD Then
            Call InitJupiterPrintServiceLevelDropdown()
        End If
        If Not IsNothing(gdtBasket) Then
            If ValidBasket() Then
                Page.Validate()
                If Page.IsValid Then
                    If CheckBasketForCalendarManagedItems() Then
                        Call ShowCalendarManagedPanel(bSkipClearDateSelection:=False)
                        Exit Sub
                    End If
                    PaintSessionVariables()
                    Call CheckOrder()
                    Try
                        ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedIndex = 0   ' kludge to quickly handle this control not being in scope
                    Catch
                    End Try
                End If
            End If
        Else
            lblError.Text = "Your basket is empty - add items to your basket before proceeding to checkout"
        End If
    End Sub
    
    Protected Sub CheckOrder()
        Call GetBasketFromSession()  ' CN new
        If Not IsNothing(gdtBasket) Then
            If ValidBasket() Then
                If Page.IsValid Then
                    If CheckBasketForCalendarManagedItems() Then
                        Call ShowCalendarManagedPanel(bSkipClearDateSelection:=False)
                        Exit Sub
                    End If

                    ' CustomLetter - what tests do we need here?
                    
                    Call PrepareAuthorisations()  ' Could probably bypass this by testing if authorisation enabled for this site
                    If gdtProductAuthorisationRequired.Rows.Count > 0 Then
                        Call ShowAuthorisationPanel()
                    Else
                        If gdtBasket.Rows.Count > 0 Then
                            Call GetAddressBookPermissions()
                            'Call ShowSelectAddressPanel()
                            Call ShowDeliveryAddressPanel()
                        Else
                            Call ShowProductList()
                            'lblProductList.Text = "Product List"
                            Call WebMsgBox.Show("An authorisation request has been sent. These items have been removed from your basket. Your basket is now empty.")
                        End If
                    End If
                End If
            End If
        Else
            lblError.Text = "Your basket is empty - add items to your basket before proceeding to checkout"
        End If
    End Sub

    Protected Sub lnkbtnRequestAuthBackToProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblError.Text = ""
        Call ShowProductList()
    End Sub

    Protected Sub lnkbtnRequestAuthShowBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lblError.Text = ""
        Call ShowBasket()
    End Sub

    Protected Sub btnFindAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call FindAddress()
        lbLookupResults.Focus()
    End Sub
    
    Protected Sub FindAddress()
        tbPostCodeLookup.Text = tbPostCodeLookup.Text.Trim.ToUpper

        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult

        objInterimResults = objLookup.ByPostcode(tbPostCodeLookup.Text, ACCOUNT_CODE, LICENSE_KEY, "")
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
            trPostCodeLookupOutput.Visible = False
        Else
            lblLookupError.Visible = False
            lbLookupResults.Visible = True
            'lblSelectADestination.Visible = True

            lbLookupResults.Items.Clear()

            If Not objInterimResults.Results Is Nothing Then
                For Each objInterimResult In objInterimResults.Results
                    lbLookupResults.Items.Add(New  _
                         ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in FindAddress logging lookup")
            End If
            trPostCodeLookupOutput.Visible = True
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
    
    Protected Sub CreateMultipleBookingStructure()
        gdtMultiAddressBooking = New DataTable
        gdtMultiAddressBooking.Columns.Add(New DataColumn("ProductKey", GetType(Long)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("Qty", GetType(Long)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("ServiceLevel", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeName", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeAddr1", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeAddr2", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeAddr3", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeTown", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeCounty", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneePostCode", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeCountryCode", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeCtcName", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneePhone", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CneeEmail", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("AddressIndex", GetType(String)))

        gdtMultiAddressBooking.Columns.Add(New DataColumn("CostCentre", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("CustRef", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("SpecialInstructions", GetType(String)))
        gdtMultiAddressBooking.Columns.Add(New DataColumn("ShippingNote", GetType(String)))
    End Sub

    Protected Sub CaptureMultipleBookingInstructions()
        Call CreateMultipleBookingStructure()
        Dim nFixedFieldsStartIndex As Integer, nIndex As Integer
        Dim nAddressIndex As Integer = 0
        Dim drMultipleBooking As DataRow
        Dim sServiceLevel As String, sCostCentre As String, sCneeName As String, sCneeAddr1 As String, sCneeAddr2 As String, sCneeAddr3 As String
        Dim sCneeTown As String, sCneeCounty As String, sCneePostCode As String, sCneeCountryCode As String
        Dim sCneeCtcName As String, sCneePhone As String, sCneeEmail As String

        Session("SB_BookingRef1") = String.Empty
        Session("SB_BookingRef2") = String.Empty
        
        lblConsignor2.Text = sBuildConsignorAddress()

        Call GetBasketFromSession()
        Dim nBasketCount As Integer = gdtBasket.Rows.Count

        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                nAddressIndex += 1
                nFixedFieldsStartIndex = gvr.Cells.Count - MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT
                For i As Integer = 0 To nFixedFieldsStartIndex
                    If i = nBasketCount Then
                        Exit For
                    End If
                    Dim tc As TableCell = gvr.Cells(i)
                    For Each c As Control In tc.Controls
                        If TypeOf c Is TextBox Then
                            Dim tb As TextBox = c
                            Dim sValue As String = tb.Text.Trim
                            If sValue.Length > 0 AndAlso IsNumeric(sValue) AndAlso CInt(sValue) > 0 Then
                                drMultipleBooking = gdtMultiAddressBooking.NewRow
                                drMultipleBooking("AddressIndex") = nAddressIndex.ToString
                                drMultipleBooking("ProductKey") = gdtBasket.Rows(i).Item("ProductKey")
                                drMultipleBooking("Qty") = sValue
                                gdtMultiAddressBooking.Rows.Add(drMultipleBooking)
                            End If
                        End If
                    Next
                Next
                nIndex = nFixedFieldsStartIndex
                Dim ddl As DropDownList = gvr.Cells(nIndex).FindControl("ddlServiceLevel")
                sServiceLevel = ddl.SelectedItem.Text
                nIndex += 1
                Dim tbCostCentre As TextBox = gvr.Cells(nIndex).FindControl("tbDistributionListCostCentre")
                sCostCentre = tbCostCentre.Text
                nIndex += 1
                Dim lblCneeName As Label = gvr.Cells(nIndex).FindControl("lblCneeName")
                sCneeName = lblCneeName.Text
                Dim lblCneeAddr1 As Label = gvr.Cells(nIndex).FindControl("lblCneeAddr1")
                sCneeAddr1 = lblCneeAddr1.Text
                Dim hidCneeAddr2 As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeAddr2")
                sCneeAddr2 = hidCneeAddr2.Value
                Dim hidCneeAddr3 As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeAddr3")
                sCneeAddr3 = hidCneeAddr3.Value
                Dim hidCneeTown As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeTown")
                sCneeTown = hidCneeTown.Value
                Dim hidCneeCounty As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeCounty")
                sCneeCounty = hidCneeCounty.Value
                Dim hidCneePostCode As HiddenField = gvr.Cells(nIndex).FindControl("hidCneePostCode")
                sCneePostCode = hidCneePostCode.Value
                Dim hidCneeCountryCode As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeCountryCode")
                sCneeCountryCode = hidCneeCountryCode.Value
                Dim hidCneeCtcName As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeCtcName")
                sCneeCtcName = hidCneeCtcName.Value
                Dim hidCneePhone As HiddenField = gvr.Cells(nIndex).FindControl("hidCneePhone")
                sCneePhone = hidCneePhone.Value
                Dim hidCneeEmail As HiddenField = gvr.Cells(nIndex).FindControl("hidCneeEmail")
                sCneeEmail = hidCneeEmail.Value
                Dim hidCustomerReference As HiddenField = gvr.Cells(nIndex).FindControl("hidCustomerReference")
                Dim hidSpecialInstructions As HiddenField = gvr.Cells(nIndex).FindControl("hidSpecialInstructions")
                Dim hidPackingNote As HiddenField = gvr.Cells(nIndex).FindControl("hidPackingNote")

                For Each dr As DataRow In gdtMultiAddressBooking.Rows
                    If CInt(dr("AddressIndex")) = nAddressIndex Then
                        dr("ServiceLevel") = sServiceLevel
                        dr("CneeName") = sCneeName
                        dr("CneeAddr1") = sCneeAddr1
                        dr("CneeAddr2") = sCneeAddr2
                        dr("CneeAddr3") = sCneeAddr3
                        dr("CneeTown") = sCneeTown
                        dr("CneeCounty") = sCneeCounty
                        dr("CneePostCode") = sCneePostCode
                        dr("CneeCountryCode") = sCneeCountryCode
                        dr("CneeCtcName") = sCneeCtcName
                        dr("CneePhone") = sCneePhone
                        dr("CneeEmail") = sCneeEmail
                        dr("CostCentre") = sCostCentre
                        dr("CustRef") = hidCustomerReference.Value
                        dr("SpecialInstructions") = hidSpecialInstructions.Value
                        dr("ShippingNote") = hidPackingNote.Value
                    End If
                Next
            End If
        Next
        ViewState("MultipleBooking") = gdtMultiAddressBooking
    End Sub
    
    Protected Sub SummariseMultipleAddressBooking()
        Dim oDataTable As DataTable = GetDistributionList(ddlDistributionList.SelectedItem.Text)
        Dim sAddressMarker As String = String.Empty
        Dim sbLine As New StringBuilder

        Call WriteSummaryTableRow1()
        For Each drMultipleBooking As DataRow In gdtMultiAddressBooking.Rows
            If drMultipleBooking("AddressIndex") <> sAddressMarker Then
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeCtcName")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeName")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeAddr1")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeAddr2")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeAddr3")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeTown")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneeCounty")))
                sbLine.Append(AddNonBlankAddressLine(drMultipleBooking("CneePostCode")))
                sbLine.Append("SERVICE LEVEL: " & drMultipleBooking("ServiceLevel"))
                sbLine.Append("<br />COST CENTRE: " & drMultipleBooking("CostCentre"))
                If drMultipleBooking("CustRef").ToString.Length > 0 Then
                    sbLine.Append("<br />CUST REF: " & drMultipleBooking("CustRef"))
                End If
                If drMultipleBooking("SpecialInstructions").ToString.Length > 0 Then
                    sbLine.Append("<br />SPEC INSTRUC: " & drMultipleBooking("SpecialInstructions"))
                End If
                If drMultipleBooking("ShippingNote").ToString.Length > 0 Then
                    sbLine.Append("<br />PACKING NOTE TEXT: " & drMultipleBooking("ShippingNote"))
                End If
                Call WriteHorizontalRule()
                Call WriteSummaryTableRowCol3(sbLine.ToString)
                sbLine.Length = 0
                sAddressMarker = drMultipleBooking("AddressIndex")
            End If
            Dim sQty As String = drMultipleBooking("Qty")
            If sQty.Length > 0 AndAlso CInt(sQty) > 0 Then
                Dim oDataView As DataView = New DataView(gdtBasket)
                oDataView.Sort = "ProductKey"
                Dim nKey As Integer = oDataView.Find(drMultipleBooking("ProductKey"))
                sbLine.Append(oDataView(nKey).Item("ProductCode"))
                sbLine.Append(" - ")
                If oDataView(nKey).Item("ProductDate").ToString.Trim <> String.Empty Then
                    sbLine.Append(oDataView(nKey).Item("ProductDate"))
                    sbLine.Append(" - ")
                End If
                sbLine.Append(oDataView(nKey).Item("Description"))
                sbLine.Append(" - QTY: ")
                sbLine.Append(drMultipleBooking("Qty"))
                sbLine.Append("<br />")
                WriteSummaryTableRowCol2(sbLine.ToString)
                sbLine.Length = 0
            End If
        Next
    End Sub
    
    Protected Function AddNonBlankAddressLine(ByVal sLine As String) As String
        sLine = sLine.Trim
        If sLine <> String.Empty Then
            AddNonBlankAddressLine = sLine & "<br />"
        Else
            AddNonBlankAddressLine = String.Empty
        End If
    End Function

    Protected Sub WriteHorizontalRule()
        Dim tr As New HtmlTableRow
        Dim tc As New HtmlTableCell
        Dim lit As New Literal

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        lit.Text = "<hr />"
        tc.Controls.Add(lit)
        tc.ColSpan = tabMultipleAddressSummary.Rows(0).Cells.Count - 2
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tabMultipleAddressSummary.Rows.Add(tr)
    End Sub
    
    Protected Sub WriteSummaryTableRow1()
        Dim tr As New HtmlTableRow
        Dim tc As HtmlTableCell

        tc = New HtmlTableCell
        tc.Width = "5%"
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tc.Width = "10%"
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tc.Width = "45%"
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tc.Width = "30%"
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tc.Width = "5%"
        tr.Cells.Add(tc)

        tabMultipleAddressSummary.Rows.Add(tr)
    End Sub
    
    Protected Sub WriteSummaryTableRowCol2(ByVal sText As String)
        Dim tr As New HtmlTableRow
        Dim lbl As Label
        Dim tc As HtmlTableCell

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        lbl = New Label
        lbl.Text = sText
        lbl.ForeColor = Navy
        tc.Controls.Add(lbl)
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tabMultipleAddressSummary.Rows.Add(tr)
    End Sub
    
    Protected Sub WriteSummaryTableRowCol3(ByVal sText As String)
        Dim tr As New HtmlTableRow
        Dim lbl As Label
        Dim tc As HtmlTableCell

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        lbl = New Label
        lbl.Text = sText
        lbl.ForeColor = Red
        tc.Controls.Add(lbl)
        tr.Cells.Add(tc)

        tc = New HtmlTableCell
        tr.Cells.Add(tc)

        tabMultipleAddressSummary.Rows.Add(tr)
    End Sub
    
    Protected Function bSufficientStock() As Boolean
        Dim nOrderTotal As Integer
        bSufficientStock = True
        For Each drBasket As DataRow In gdtBasket.Rows
            nOrderTotal = 0
            For Each drMultipleBooking As DataRow In gdtMultiAddressBooking.Rows
                If drBasket("ProductKey") = drMultipleBooking("ProductKey") AndAlso IsNumeric(drMultipleBooking("Qty")) Then
                    nOrderTotal = nOrderTotal + CInt(drMultipleBooking("Qty"))
                End If
            Next
            If nOrderTotal > CInt(drBasket("QtyAvailable")) Then
                WebMsgBox.Show("Total quantity (" & nOrderTotal.ToString & ") ordered for item " & drBasket("ProductCode") & " exceeds the quantity available (" & drBasket("QtyAvailable") & ") - please adjust the order quantities.")
                bSufficientStock = False
            End If
        Next
    End Function
    
    Protected Sub btnMultipleAddressOrderFinalConfirmation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PerformMultipleAddressOrder()
    End Sub
    
    Protected Sub PerformMultipleAddressOrder()
        Call MultipleAddressOrderFinalConfirmation()
    End Sub
    
    Protected Sub MultipleAddressOrderFinalConfirmation()
        If bValidMultipleAddressQuantities() AndAlso bValidCostCentres() Then
            Call CaptureMultipleBookingInstructions()
            If bOrdersToProcess() Then
                If bSufficientStock() Then
                    Call SummariseMultipleAddressBooking()
                    ViewState("SB_MultipleBooking") = gdtMultiAddressBooking
                    ShowCompleteMultipleAddressBookingPanel()
                Else
                    WebMsgBox.Show("There is insufficient stock available to fulfil one or more of your orders. Please adjust your order quantities.")
                End If
            Else
                WebMsgBox.Show("You have not specified any orders.")
            End If
        End If
    End Sub
    
    Protected Function bQuantitiesZeroOrBlank(ByVal gvr As GridViewRow) As Boolean
        Dim nFixedFieldsStartIndex As Integer
        Dim nBasketCount As Integer = gdtBasket.Rows.Count
        bQuantitiesZeroOrBlank = False
        nFixedFieldsStartIndex = gvr.Cells.Count - MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT
        For i As Integer = 0 To nFixedFieldsStartIndex
            If i = nBasketCount Then
                Exit For
            End If
            Dim tc As TableCell = gvr.Cells(i)
            For Each c As Control In tc.Controls
                If TypeOf c Is TextBox Then
                    Dim tb As TextBox = c
                    tb.Text = tb.Text.Trim
                    If tb.Text = String.Empty Then
                    Else
                        If IsNumeric(tb.Text) AndAlso CInt(tb.Text) = 0 Then
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
            Next
        Next
        bQuantitiesZeroOrBlank = True
    End Function
    
    Protected Function bValidCostCentres() As Boolean
        Dim tc As TableCell
        Dim tb As TextBox
        bValidCostCentres = True
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                tc = gvr.Cells(MULTIPLE_ADDRESS_COST_CENTRE_COLUMN - 1)
                tb = tc.FindControl("tbDistributionListCostCentre")
                tb.Text = tb.Text.Trim
                If Not bQuantitiesZeroOrBlank(gvr) AndAlso tb.Text = String.Empty Then
                    bValidCostCentres = False
                    WebMsgBox.Show("One or more cost centres missing - you must specify a cost centre for each destination that will receive a delivery.")
                    Exit For
                End If
            End If
        Next
    End Function
    
    Protected Sub btnSetAllQuantities_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbSetAllQuantities.Text = tbSetAllQuantities.Text.Trim
        If IsInteger(tbSetAllQuantities.Text) Then
            Call SetAllQuantities(tbSetAllQuantities.Text)
        Else
            WebMsgBox.Show("Please enter a number (zero or greater)")
        End If
    End Sub
    
    Protected Sub SetAllQuantities(ByVal sValue As String)
        Dim nFixedFieldsStartIndex As Integer
        Call GetBasketFromSession()
        Dim nBasketCount As Integer = gdtBasket.Rows.Count
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                nFixedFieldsStartIndex = gvr.Cells.Count - MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT
                For i As Integer = 0 To nFixedFieldsStartIndex
                    If i = nBasketCount Then
                        Exit For
                    End If
                    Dim tc As TableCell = gvr.Cells(i)
                    For Each c As Control In tc.Controls
                        If TypeOf c Is TextBox Then
                            Dim tb As TextBox = c
                            tb.Text = tbSetAllQuantities.Text
                        End If
                    Next
                Next
            End If
        Next
    End Sub
    
    Protected Sub lnkbtnClearAllQuantities_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetAllQuantities("")
    End Sub
    
    Protected Sub SetAllCostCentres(ByVal sCostCentre As String)
        Dim tb As TextBox
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                tb = gvr.Cells(MULTIPLE_ADDRESS_COST_CENTRE_COLUMN - 1).FindControl("tbDistributionListCostCentre")
                tb.Text = sCostCentre
            End If
        Next
    End Sub
    
    Protected Sub btnSetAllCostCentres_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetAllCostCentres(tbSetAllCostCentres.Text)
    End Sub

    Protected Sub lnkbtnClearAllCostCentres_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetAllCostCentres("")
    End Sub

    Protected Function bValidMultipleAddressQuantities() As Boolean
        Dim nFixedFieldsStartIndex As Integer
        bValidMultipleAddressQuantities = True
        Dim nLineCount As Integer = 0
        Call GetBasketFromSession()
        Dim nBasketCount As Integer = gdtBasket.Rows.Count
        For Each gvr As GridViewRow In gvDistributionList.Rows
            nLineCount += 1
            If gvr.RowType = DataControlRowType.DataRow Then
                nFixedFieldsStartIndex = gvr.Cells.Count - MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT
                For i As Integer = 0 To nFixedFieldsStartIndex
                    If i = nBasketCount Then
                        Exit For
                    End If
                    Dim tc As TableCell = gvr.Cells(i)
                    For Each c As Control In tc.Controls
                        If TypeOf c Is TextBox Then
                            Dim tb As TextBox = c
                            tb.Text = tb.Text.Trim
                            If Not (tb.Text = "" Or IsInteger(tb.Text)) Then
                                WebMsgBox.Show("One of the quantities on line " & nLineCount.ToString & " is neither a valid number nor blank - please correct")
                                bValidMultipleAddressQuantities = False
                                Exit Function
                            End If
                        End If
                    Next
                Next
            End If
        Next
    End Function
    
    Protected Function bOrdersToProcess() As Boolean
        Dim nFixedFieldsStartIndex As Integer
        bOrdersToProcess = False
        Call GetBasketFromSession()
        Dim nBasketCount As Integer = gdtBasket.Rows.Count
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                nFixedFieldsStartIndex = gvr.Cells.Count - MULTIPLE_ADDRESS_FIXED_FIELDS_COUNT
                For i As Integer = 0 To nFixedFieldsStartIndex
                    If i = nBasketCount Then
                        Exit For
                    End If
                    Dim tc As TableCell = gvr.Cells(i)
                    For Each c As Control In tc.Controls
                        If TypeOf c Is TextBox Then
                            Dim tb As TextBox = c
                            tb.Text = tb.Text.Trim
                            If tb.Text <> String.Empty AndAlso CInt(tb.Text) > 0 Then
                                bOrdersToProcess = True
                                Exit Function
                            End If
                        End If
                    Next
                Next
            End If
        Next
    End Function
    
    Protected Sub btnConfirmMultipleAddressOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ConfirmMultipleAddressOrder()
    End Sub
    
    Protected Sub ConfirmMultipleAddressOrder()
        gdtMultiAddressBooking = ViewState("SB_MultipleBooking")
        Call SubmitConsignments()
        Call ShowBookingConfirmationPanel()
        'Call UpdateAuthorisations()
        Call ClearOrderSessionVariables()
    End Sub

    Protected Function nGetHighestAddressIndex() As Integer
        Dim nAddressIndex As Integer = 0
        For Each drMultipleBooking As DataRow In gdtMultiAddressBooking.Rows
            If CInt(drMultipleBooking("AddressIndex")) > nAddressIndex Then
                nAddressIndex = CInt(drMultipleBooking("AddressIndex"))
            End If
        Next
        nGetHighestAddressIndex = nAddressIndex
    End Function
    
    Protected Sub SetCommonConsignmentValues()
        Session("SB_BookingRef1") = ""
        Session("SB_BookingRef2") = ""
    End Sub
    
    Protected Sub SetConsigneeConsignmentValues(ByVal drMultipleBooking As DataRow)
        Session("SB_CneeCompany") = drMultipleBooking("CneeName")
        Session("SB_CneeAddr1") = drMultipleBooking("CneeAddr1")
        Session("SB_CneeAddr2") = drMultipleBooking("CneeAddr2")
        Session("SB_CneeAddr3") = drMultipleBooking("CneeAddr3")
        Session("SB_CneeTown") = drMultipleBooking("CneeTown")
        Session("SB_CneeState") = drMultipleBooking("CneeCounty")
        Session("SB_CneePostCode") = drMultipleBooking("CneePostCode")
        Session("SB_CneeCountryKey") = drMultipleBooking("CneeCountryCode")
        Session("SB_CneeCountryName") = "UK"
        Session("SB_CneeCtcName") = drMultipleBooking("CneeCtcName")
        Session("SB_CneeCtcTel") = drMultipleBooking("CneePhone")
        Session("SB_CneeCtcEmail") = drMultipleBooking("CneeEmail")
        Session("SB_BookingRef3") = drMultipleBooking("CostCentre")
        Session("SB_BookingRef4") = drMultipleBooking("CustRef")
        If drMultipleBooking("SpecialInstructions").ToString.Length > 0 Then
            Session("SB_SpecialInstructions") = "SERVICE LEVEL: " & drMultipleBooking("ServiceLevel") & ";" & drMultipleBooking("SpecialInstructions")
        Else
            Session("SB_SpecialInstructions") = "SERVICE LEVEL: " & drMultipleBooking("ServiceLevel")
        End If
        Session("SB_ShippingNote") = drMultipleBooking("ShippingNote")
    End Sub
    
    Protected Sub SubmitConsignments()
        Dim oDataTable As DataTable = GetDistributionList(ddlDistributionList.SelectedItem.Text)
        Dim sAddressMarker As String = String.Empty
        Dim gnrdicOrderItems As New Dictionary(Of Integer, Integer)
        Dim gnrclstConsignmentNumbers As New List(Of Integer)
        Dim bIsSetConsigneeConsignmentValues As Boolean
        Call SetCommonConsignmentValues()
        For i As Integer = 1 To nGetHighestAddressIndex()
            gnrdicOrderItems.Clear()
            bIsSetConsigneeConsignmentValues = False
            For Each drMultiAddressBooking As DataRow In gdtMultiAddressBooking.Rows
                If i = drMultiAddressBooking("AddressIndex") Then
                    If Not bIsSetConsigneeConsignmentValues Then
                        Call SetConsigneeConsignmentValues(drMultiAddressBooking)
                        bIsSetConsigneeConsignmentValues = True
                    End If
                    Dim sQty As String = drMultiAddressBooking("Qty").ToString.Trim
                    If sQty <> String.Empty AndAlso CInt(sQty) > 0 Then
                        gnrdicOrderItems.Add(CInt(drMultiAddressBooking("ProductKey")), CInt(sQty))
                    End If
                End If
            Next
            If gnrdicOrderItems.Count > 0 Then
                If Not pbAuthorisationRequired Then
                    gnrclstConsignmentNumbers.Add(nSubmitConsignment(gnrdicOrderItems))
                    ' Call UpdateAuthorisations() ?
                Else
                    gnrclstConsignmentNumbers.Add(PlaceOrderOnHold(gnrdicOrderItems))
                End If
            End If
        Next
        ' for queued orders need to do equivalent of Call ShowBookingQueuedConfirmationPanel()
        For Each nConsignmentNumber As Integer In gnrclstConsignmentNumbers
            lblConsignmentNo.Text = lblConsignmentNo.Text & nConsignmentNumber.ToString & " "
        Next
    End Sub

    Protected Function nSubmitConsignment(ByVal dicOrderItems As Dictionary(Of Integer, Integer)) As Integer
        Dim sSpecialInstr As String
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CLng(Session("UserKey"))
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CLng(Session("CustomerKey"))
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = Session("SB_BookingRef1")
        oCmdAddBooking.Parameters.Add(param3)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = Session("SB_BookingRef2")
        oCmdAddBooking.Parameters.Add(param4)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = Session("SB_BookingRef3")
        oCmdAddBooking.Parameters.Add(param5)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = Session("SB_BookingRef4")
        oCmdAddBooking.Parameters.Add(param6)
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        sSpecialInstr = Session("SB_SpecialInstructions")
        sSpecialInstr = Replace(sSpecialInstr, vbCrLf, " ")
        
        If pbBasketContainsWURSCriticalProducts Then
            sSpecialInstr += "SYSMSG: CRITICAL WURS ORDER - SEND BY COURIER "
        End If

        If IsVSOE() AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery00.Checked Then
            sSpecialInstr = "[STANDARD DELIVERY] " & sSpecialInstr
        End If
        
        If IsOlympus() Then
            'sSpecialInstr = "SYSMSG: PLEASE POST 2ND CLASS " & sSpecialInstr
        End If

        If IsRioTinto() Then
            sSpecialInstr = "SYSMSG: PLEASE POST UK ITEMS UNDER 2KG " & sSpecialInstr
        End If

        param7.Value = sSpecialInstr
        oCmdAddBooking.Parameters.Add(param7)
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = Session("SB_ShippingNote")
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
        param13.Value = psCnorCompany
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = psCnorAddr1
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = psCnorAddr2
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = psCnorAddr3
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = psCnorTown
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = psCnorState
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = psCnorPostCode
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = CLng(psCnorCountryKey)
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = psCnorCtcName
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = psCnorCtcTel
        oCmdAddBooking.Parameters.Add(param22)
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = psCnorCtcEmail
        oCmdAddBooking.Parameters.Add(param23)
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = Session("SB_CneeCompany")
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = Session("SB_CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = Session("SB_CneeAddr2")
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = Session("SB_CneeAddr3")
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = Session("SB_CneeTown")
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = Session("SB_CneeState")
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = Session("SB_CneePostCode")
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = Session("SB_CneeCountryKey")
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = Session("SB_CneeCtcName")
        oCmdAddBooking.Parameters.Add(param33)
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = Session("SB_CneeCtcTel")
        oCmdAddBooking.Parameters.Add(param34)
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = Session("SB_CneeCtcEmail")
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
                For Each kvp As KeyValuePair(Of Integer, Integer) In dicOrderItems
                    Dim lProductKey As Long = kvp.Key
                    Dim lPickQuantity As Long = kvp.Value
                    Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                    oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                    Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                    param51.Value = CLng(Session("UserKey"))
                    oCmdAddStockItem.Parameters.Add(param51)
                    Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                    param52.Value = CLng(Session("CustomerKey"))
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
                lblError.Text = "Error adding Web Booking [BookingKey=0]."
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                nSubmitConsignment = lConsignmentKey
                If gnMode = MODE_JUPITER_POD Then
                    Call NotifyJupiterPDFOrder(lConsignmentKey)
                    Call ExecuteQueryToDataTable("UPDATE Consignment SET AgentRef = '1. ORDER PLACED, AWAITING PRINTER RESPONSE' + ' (' + CAST(REPLACE(CONVERT(VARCHAR(11),  GETDATE(), 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), GETDATE(), 108)),1,5) + ')', AgentAWB = '' WHERE [key] = " & lConsignmentKey)
                End If
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
    End Function

    Protected Sub NotifyJupiterPDFOrder(ByVal nConsignmentKey As Int32)
        Dim sSQL As String = "SELECT EmailAddr FROM ClientData_Jupiter_EventNotification WHERE EventCode LIKE '%order%'"
        Dim sbMessage As New StringBuilder
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        sbMessage.Append("JUPITER ASSET MANAGEMENT - Print Order Notification")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("A new order for printed materials has been placed.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Order #:")
        sbMessage.Append(nConsignmentKey.ToString)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Visit http://my.transworld.eu.com/jupiter to view orders.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Please do not reply to this email as replies are not monitored.  Thank you.")
        sbMessage.Append(Environment.NewLine)
        sbMessage.Append("Transworld")
        Dim sPlainTextBody As String = sbMessage.ToString
        Dim sHTMLBody As String = sbMessage.ToString.Replace(Environment.NewLine, "<br />" & Environment.NewLine)
        For Each dr As DataRow In dt.Rows
            Call SendMail("JUPITER_ORDER_ALERT", dr(0), "JUPITER Print Order Alert", sPlainTextBody, sHTMLBody)
        Next
    End Sub
    
    Protected Function PlaceOrderOnHold(ByVal dicOrderItems As Dictionary(Of Integer, Integer)) As Integer
        Dim lHoldingQueueKey As Long
        Dim guidAuthorisationGUID As Guid
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderPlaceOnHold", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure

        guidAuthorisationGUID = Guid.NewGuid
        Dim paramAuthorisationGUID As SqlParameter = New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20)
        paramAuthorisationGUID.Value = guidAuthorisationGUID.ToString
        oCmdAddBooking.Parameters.Add(paramAuthorisationGUID)

        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CLng(Session("UserKey"))
        oCmdAddBooking.Parameters.Add(param1)

        Dim paramAuthoriserKey As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int, 4)
        paramAuthoriserKey.Value = pnAuthoriser
        oCmdAddBooking.Parameters.Add(paramAuthoriserKey)

        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CLng(Session("CustomerKey"))
        oCmdAddBooking.Parameters.Add(param2)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = Session("SB_BookingRef1")
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = Session("SB_BookingRef2")
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = Session("SB_BookingRef3")
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = Session("SB_BookingRef4")
        oCmdAddBooking.Parameters.Add(param6)
        
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param6a)

        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = Session("SB_SpecialInstructions").ToString.Replace(vbCrLf, " ")
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = Session("SB_ShippingNote")
        oCmdAddBooking.Parameters.Add(param8)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = psCnorCompany
        oCmdAddBooking.Parameters.Add(param13)

        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = psCnorAddr1
        oCmdAddBooking.Parameters.Add(param14)

        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = psCnorAddr2
        oCmdAddBooking.Parameters.Add(param15)

        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = psCnorAddr3
        oCmdAddBooking.Parameters.Add(param16)

        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = psCnorTown
        oCmdAddBooking.Parameters.Add(param17)

        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = psCnorState
        oCmdAddBooking.Parameters.Add(param18)

        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = psCnorPostCode
        oCmdAddBooking.Parameters.Add(param19)

        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = CLng(psCnorCountryKey)
        oCmdAddBooking.Parameters.Add(param20)

        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = psCnorCtcName
        oCmdAddBooking.Parameters.Add(param21)

        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = psCnorCtcTel
        oCmdAddBooking.Parameters.Add(param22)

        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = psCnorCtcEmail
        oCmdAddBooking.Parameters.Add(param23)

        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = Session("SB_CneeCompany")
        oCmdAddBooking.Parameters.Add(param25)

        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = Session("SB_CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)

        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = Session("SB_CneeAddr2")
        oCmdAddBooking.Parameters.Add(param27)

        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = Session("SB_CneeAddr3")
        oCmdAddBooking.Parameters.Add(param28)

        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = Session("SB_CneeTown")
        oCmdAddBooking.Parameters.Add(param29)

        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = Session("SB_CneeState")
        oCmdAddBooking.Parameters.Add(param30)

        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = Session("SB_CneePostCode")
        oCmdAddBooking.Parameters.Add(param31)

        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = Session("SB_CneeCountryKey")
        oCmdAddBooking.Parameters.Add(param32)

        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = Session("SB_CneeCtcName")
        oCmdAddBooking.Parameters.Add(param33)

        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = Session("SB_CneeCtcTel")
        oCmdAddBooking.Parameters.Add(param34)

        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = Session("SB_CneeCtcEmail")
        oCmdAddBooking.Parameters.Add(param35)

        Dim param36 As SqlParameter = New SqlParameter("@MsgToAuthoriser", SqlDbType.NVarChar, 1000)
        param36.Value = tbAuthoriserMessageSingleAddressOrder.Text.Replace(Environment.NewLine, " ")
        oCmdAddBooking.Parameters.Add(param36)

        Dim param37 As SqlParameter = New SqlParameter("@HoldingQueueKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)

        Try
            oConn.Open()
            oCmdAddBooking.ExecuteNonQuery()
            lHoldingQueueKey = CLng(oCmdAddBooking.Parameters("@HoldingQueueKey").Value.ToString)
            If lHoldingQueueKey > 0 Then
                For Each kvp As KeyValuePair(Of Integer, Integer) In dicOrderItems
                    Try
                        Dim lProductKey As Long = kvp.Key
                        Dim lPickQuantity As Long = kvp.Value
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderHoldingQueueItemAdd", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param53 As SqlParameter = New SqlParameter("@OrderHoldingQueueKey", SqlDbType.Int, 4)
                        param53.Value = lHoldingQueueKey
                        oCmdAddStockItem.Parameters.Add(param53)
                        Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param54.Value = lProductKey
                        oCmdAddStockItem.Parameters.Add(param54)
                        Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                        param56.Value = lPickQuantity
                        oCmdAddStockItem.Parameters.Add(param56)
                        oCmdAddStockItem.Connection = oConn
                        oCmdAddStockItem.ExecuteNonQuery()
                    Catch ex As Exception
                        NotifyException("PlaceOrderOnHold", "Could not add product to holding queue", ex)
                    End Try
                Next
                EmailAuthoriser(guidAuthorisationGUID.ToString)
            Else
                NotifyException("PlaceOrderOnHold", "Internal error - no product selected")
            End If
        Catch ex As Exception
            NotifyException("PlaceOrderOnHold", "Could not add order to holding queue", ex)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Function GetBookingRefs() As BookingRefs
        Dim oBookingRefs As New BookingRefs
        oBookingRefs._BookingRef1 = Session("SB_BookingRef1")
        oBookingRefs._BookingRef2 = Session("SB_BookingRef2")
        oBookingRefs._BookingRef3 = Session("SB_BookingRef3")
        oBookingRefs._BookingRef4 = Session("SB_BookingRef4")
        'End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_1_BLACKROCK Then
            oBookingRefs._BookingRef1 = String.Empty
            oBookingRefs._BookingRef2 = String.Empty
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration1CostCentre.Text
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_2_LEGACY_SINGLE_MANDATORY_CUSTREF3 Or _
              plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Or _
                plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_16_VSOE Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration2AdditionalRefB.Text.Trim
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration2AdditionalRefC.Text.Trim
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration2BookingRef.Text.Trim
            oBookingRefs._BookingRef4 = tbPerCustomerConfiguration2AdditionalRefA.Text.Trim
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_3_LEGACY_SINGLE_MANDATORY_UNPROMPTED_COST_CENTRE Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration3AdditionalRefB.Text.Trim
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration3AdditionalRefC.Text.Trim
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration3CostCentre.Text.Trim
            oBookingRefs._BookingRef4 = tbPerCustomerConfiguration3AdditionalRefA.Text.Trim
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_5_KODAK Then
            oBookingRefs._BookingRef1 = String.Empty
            oBookingRefs._BookingRef2 = String.Empty
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration5CostCentre.Text
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_6_CIMA Then
            'oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration6Department.SelectedItem.Text
            'oBookingRefs._BookingRef2 = Session("UserName")
            'oBookingRefs._BookingRef3 = tbPerCustomerConfiguration6CostCentre.Text.Trim
            'oBookingRefs._BookingRef4 = tbPerCustomerConfiguration6Reference.Text.Trim
            oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration6Department.SelectedItem.Text
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration6CostCentre.Text.Trim
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration6Reference.Text.Trim
            oBookingRefs._BookingRef4 = Session("UserName")
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
            oBookingRefs._BookingRef1 = Session("BookingRef")
            oBookingRefs._BookingRef2 = Session("Rating")
            oBookingRefs._BookingRef3 = String.Empty
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_12_ATKINS Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration12CostCentre.Text.Trim
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration12AdditionalRefA.Text.Trim
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration12AdditionalRefB.Text.Trim
            oBookingRefs._BookingRef4 = tbPerCustomerConfiguration12AdditionalRefC.Text.Trim
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_19_INSIGHT Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration19CostCentre.Text
            oBookingRefs._BookingRef2 = String.Empty
            oBookingRefs._BookingRef3 = String.Empty
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_20_DAT Then
            oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration20CostCentre.SelectedItem.Text
            oBookingRefs._BookingRef2 = String.Empty
            oBookingRefs._BookingRef3 = String.Empty
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_21_UNICRD Then
            oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration21CostCentre.SelectedItem.Text
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration21CustRef.Text
            oBookingRefs._BookingRef3 = String.Empty
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_22_RIOTINTO Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration22CostCentre.Text
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration22RequestedBy.Text
            oBookingRefs._BookingRef3 = String.Empty
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_23_ARTHRITIS Then
            oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration23CostCentre.SelectedValue
            oBookingRefs._BookingRef2 = ddlPerCustomerConfiguration23Category.SelectedValue
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration23PONumber.Text
            oBookingRefs._BookingRef4 = String.Empty
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_24_PROMOVERITAS Then
            oBookingRefs._BookingRef1 = tbPerCustomerConfiguration24Reference.Text
            oBookingRefs._BookingRef2 = tbPerCustomerConfiguration24RequestedBy.Text
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration24RecipientName.Text
            oBookingRefs._BookingRef4 = tbPerCustomerConfiguration24Products.Text
        End If

        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_29_IRWINMITCHELL Then
            oBookingRefs._BookingRef1 = ddlPerCustomerConfiguration29IrwinMitchellBudgetCode.SelectedItem.Text
            oBookingRefs._BookingRef2 = ddlPerCustomerConfiguration29IrwinMitchellDepartment.SelectedItem.Text
            oBookingRefs._BookingRef3 = tbPerCustomerConfiguration29Event.Text
            oBookingRefs._BookingRef4 = tbPerCustomerConfiguration29BDName.Text
        End If

        GetBookingRefs = oBookingRefs
    End Function
    
    Public Class BookingRefs
        Public _BookingRef1 As String
        Public _BookingRef2 As String
        Public _BookingRef3 As String
        Public _BookingRef4 As String
    End Class
    
    Protected Sub SubmitOrder() ' single consignment only; for multiple consignments from single order (eg ODP & non-ODP) see SubmitConsignments
        If IsValid Then
            Dim sSpecialInstr As String
            Dim lBookingKey As Long
            'Dim plConsignmentKey As Long
            Dim BookingFailed As Boolean
            Dim drv As DataRowView
            Dim oConn As New SqlConnection(gsConn)
            Dim oTrans As SqlTransaction
            Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
            oCmdAddBooking.CommandType = CommandType.StoredProcedure
            lblError.Text = ""
            Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
            param1.Value = CLng(Session("UserKey"))
            oCmdAddBooking.Parameters.Add(param1)
            Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
            param2.Value = CLng(Session("CustomerKey"))
            oCmdAddBooking.Parameters.Add(param2)
            Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
            param2a.Value = "WEB_BOOKING"
            oCmdAddBooking.Parameters.Add(param2a)

            Dim oBookingRefs As BookingRefs = GetBookingRefs()
            Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
            param3.Value = oBookingRefs._BookingRef1
            'If gnMode = MODE_JUPITER_POD Then
            If IsJupiter() Then
                param3.Value = ddlPerCustomerConfiguration30JupiterBudgetCode.SelectedItem.Text
            End If
            oCmdAddBooking.Parameters.Add(param3)
            Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
            param4.Value = oBookingRefs._BookingRef2
            oCmdAddBooking.Parameters.Add(param4)
            Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
            param5.Value = oBookingRefs._BookingRef3
            If gnMode = MODE_JUPITER_POD Then
                param5.Value = ddlPerCustomerConfiguration30PrintServiceLevel.SelectedItem.Text
            End If
            oCmdAddBooking.Parameters.Add(param5)
            Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
            param6.Value = oBookingRefs._BookingRef4
            If gnMode = MODE_JUPITER_POD Then
                param6.Value = lblPerCustomerConfiguration30TotalPrintCost.Text
            End If
            oCmdAddBooking.Parameters.Add(param6)
            
            Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
            param6a.Value = Nothing
            oCmdAddBooking.Parameters.Add(param6a)
            
            If IsVSOE() AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery00.Checked Then
                Session("SB_SpecialInstructions") = Session("SB_SpecialInstructions").ToString.Replace("AUTHORISED FOR EXPRESS DELIVERY", "[DENIED EXPRESS DELIVERY DUE TO ATTEMPT TO BYPASS AUTHORISATION] ")
            End If

            Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
            sSpecialInstr = Session("SB_SpecialInstructions")
            sSpecialInstr = Replace(sSpecialInstr, vbCrLf, " ")
            
            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_4_HYSTER_YALE Or _
              plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_7_OECORP Or _
                    IsVSAL() Then
                
                Dim sTemp = Session("SB_SpecialInstructions").ToString.Trim
                sTemp = Replace(sTemp, vbCrLf, " ")
                sSpecialInstr = "Service Level: " & psServiceLevel
                If sTemp <> String.Empty Then
                    sSpecialInstr = sSpecialInstr & " Instructions: " & sTemp
                End If
            End If

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
                sSpecialInstr = Session("PCID") & ". " & sSpecialInstr
            End If

            If pbBasketContainsWURSCriticalProducts Then
                sSpecialInstr += "SYSMSG: CRITICAL WURS ORDER - SEND BY COURIER"
            End If

            If IsAAT() Then
                sSpecialInstr = "[PRIORITY: " & ddlPerCustomerConfiguration17ServiceLevel.SelectedItem.Text & "] " & sSpecialInstr
            End If

            If IsVSOE() AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery00.Checked Then
                sSpecialInstr = "[STANDARD DELIVERY] " & sSpecialInstr
            End If
        
            If IsOlympus() Then
                'sSpecialInstr = "SYSMSG: PLEASE POST 2ND CLASS " & sSpecialInstr
            End If
            
            If IsRioTinto() Then
                sSpecialInstr = "SYSMSG: PLEASE POST UK ITEMS UNDER 2KG " & sSpecialInstr
            End If

            param7.Value = sSpecialInstr
            oCmdAddBooking.Parameters.Add(param7)
            Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
            param8.Value = Session("SB_ShippingNote")
            oCmdAddBooking.Parameters.Add(param8)

            If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
                Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
                param9.Value = Session("RO")
                oCmdAddBooking.Parameters.Add(param9)
            Else
                Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
                param9.Value = "STOCK ITEM"
                oCmdAddBooking.Parameters.Add(param9)
            End If

            Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
            param10.Value = -1
            oCmdAddBooking.Parameters.Add(param10)
            Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
            param11.Value = "PRINTED MATTER - FREE DOMICILE"
            oCmdAddBooking.Parameters.Add(param11)
            Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
            param13.Value = psCnorCompany
            oCmdAddBooking.Parameters.Add(param13)
            Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
            param14.Value = psCnorAddr1
            oCmdAddBooking.Parameters.Add(param14)
            Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
            param15.Value = psCnorAddr2
            oCmdAddBooking.Parameters.Add(param15)
            Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
            param16.Value = psCnorAddr3
            oCmdAddBooking.Parameters.Add(param16)
            Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
            param17.Value = psCnorTown
            oCmdAddBooking.Parameters.Add(param17)
            Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
            param18.Value = psCnorState
            oCmdAddBooking.Parameters.Add(param18)
            Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
            param19.Value = psCnorPostCode
            oCmdAddBooking.Parameters.Add(param19)
            Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
            param20.Value = CLng(psCnorCountryKey)
            oCmdAddBooking.Parameters.Add(param20)
            Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
            param21.Value = psCnorCtcName
            oCmdAddBooking.Parameters.Add(param21)
            Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
            param22.Value = psCnorCtcTel
            oCmdAddBooking.Parameters.Add(param22)
            Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
            param23.Value = psCnorCtcEmail
            oCmdAddBooking.Parameters.Add(param23)
            Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
            param24.Value = 0
            oCmdAddBooking.Parameters.Add(param24)
            Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
            param25.Value = Session("SB_CneeCompany")
            oCmdAddBooking.Parameters.Add(param25)
            Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
            param26.Value = Session("SB_CneeAddr1")
            oCmdAddBooking.Parameters.Add(param26)
            Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
            param27.Value = Session("SB_CneeAddr2")
            oCmdAddBooking.Parameters.Add(param27)
            Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
            param28.Value = Session("SB_CneeAddr3")
            oCmdAddBooking.Parameters.Add(param28)
            Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
            param29.Value = Session("SB_CneeTown")
            oCmdAddBooking.Parameters.Add(param29)
            Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
            param30.Value = Session("SB_CneeState")
            oCmdAddBooking.Parameters.Add(param30)
            Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
            param31.Value = Session("SB_CneePostCode")
            oCmdAddBooking.Parameters.Add(param31)
            Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
            param32.Value = Session("SB_CneeCountryKey")
            oCmdAddBooking.Parameters.Add(param32)
            Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
            param33.Value = Session("SB_CneeCtcName")
            oCmdAddBooking.Parameters.Add(param33)
            Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
            param34.Value = Session("SB_CneeCtcTel")
            oCmdAddBooking.Parameters.Add(param34)
            Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
            param35.Value = Session("SB_CneeCtcEmail")
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
                plConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
                If lBookingKey > 0 Then
                    Dim BasketView As New DataView
                    Call GetBasketFromSession()
                    BasketView = gdtBasket.DefaultView
                    If BasketView.Count > 0 Then
                        For Each drv In BasketView
                            Dim lProductKey As Long = CLng(drv("ProductKey"))
                            Dim lPickQuantity As Long = CLng(drv("QtyToPick"))
                            Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                            oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                            Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                            param51.Value = CLng(Session("UserKey"))
                            oCmdAddStockItem.Parameters.Add(param51)
                            Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                            param52.Value = CLng(Session("CustomerKey"))
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
                            param57.Value = plConsignmentKey
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

                        If plPerCustomerConfiguration = PER_CUSTOMER_CONFIGURATION_8_MAN Then
                            If tbPerCustomerConfiguration8MDSOrderRef.Text.Trim.Length > 0 Then
                                Dim oCmdUpdateMSDOrderRef As SqlCommand = New SqlCommand("MDS_Order_UpdateLocalEntryMDSOrderRef", oConn)
                                oCmdUpdateMSDOrderRef.CommandType = CommandType.StoredProcedure
                                Dim paramLogisticBookingKey As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                                paramLogisticBookingKey.Value = lBookingKey
                                oCmdUpdateMSDOrderRef.Parameters.Add(paramLogisticBookingKey)
                                Dim paramMDSOrderRef As SqlParameter = New SqlParameter("@MDSOrderRef", SqlDbType.NVarChar, 50)
                                paramMDSOrderRef.Value = tbPerCustomerConfiguration8MDSOrderRef.Text.Trim
                                oCmdUpdateMSDOrderRef.Parameters.Add(paramMDSOrderRef)
                                oCmdUpdateMSDOrderRef.Connection = oConn
                                oCmdUpdateMSDOrderRef.Transaction = oTrans
                                oCmdUpdateMSDOrderRef.ExecuteNonQuery()
                            End If
                        End If
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
                    lblConsignmentNo.Text = plConsignmentKey.ToString

                    If IsHysterOrYale() Then
                        Dim sbMessage As New StringBuilder
                        sbMessage.Append("Consignment number: " & lblConsignmentNo.Text & "<br />" & Environment.NewLine)
                        sbMessage.Append("Consignment placed on: " & DateTime.Now.ToShortDateString & " " & DateTime.Now.ToShortTimeString & "<br />" & Environment.NewLine)
                        sbMessage.Append("Consignment placed by: " & Session("UserName") & "<br />" & Environment.NewLine)
                        sbMessage.Append("Service level: " & psServiceLevel & "<br />" & Environment.NewLine)
                        sbMessage.Append("Estimated weight: " & lblPerCustomerConfiguration4Confirmation2Weight.Text & "<br />" & Environment.NewLine)
                        sbMessage.Append("Estimated cost: " & lblPerCustomerConfiguration4Confirmation2BasketShippingCost.Text & "<br />" & Environment.NewLine)
                        sbMessage.Append("<br />")
                        sbMessage.Append(hidCostCalculationTrace.Value)
                        Call SendMail("HYSTER_COST_ESTIMATE", "yvonne.tudgay@transworld.eu.com", "Cost estimate for Hyster or Yale order " & lblConsignmentNo.Text, sbMessage.ToString, sbMessage.ToString)
                        Dim sSQL As String = "INSERT INTO ClientData_HysterYale_ConsignmentCostEstimate (ConsignmentKey, ServiceLevel, EstimatedWeight, EstimatedCost) VALUES (" & lblConsignmentNo.Text & ", '" & psServiceLevel & "', " & lblPerCustomerConfiguration4Confirmation2Weight.Text.Replace(",", ".") & ", " & lblPerCustomerConfiguration4Confirmation2BasketShippingCost.Text.Replace(",", ".") & ")"
                        Call ExecuteQueryToDataTable(sSQL)
                    End If

                    Call ShowBookingConfirmationPanel()
                    If gnMode = MODE_JUPITER_POD Then
                        Call NotifyJupiterPDFOrder(plConsignmentKey)
                        Call ExecuteQueryToDataTable("UPDATE Consignment SET AgentRef = 'ORDER PLACED, AWAITING PRINTER RESPONSE', AgentAWB = '' WHERE [key] = " & plConsignmentKey)
                    End If

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

            txtCneeName.Text = objAddress.OrganisationName
            txtCneeAddr1.Text = objAddress.Line1
            txtCneeAddr2.Text = objAddress.Line2
            txtCneeAddr3.Text = objAddress.Line3
            txtCneeCity.Text = objAddress.PostTown
            txtCneePostCode.Text = objAddress.Postcode
            txtCneeState.Text = objAddress.County

            For i As Integer = 0 To ddlCneeCountry.Items.Count - 1
                If ddlCneeCountry.Items(i).Text = "U.K." Or ddlCneeCountry.Items(i).Text = "UK" Then
                    ddlCneeCountry.SelectedIndex = i
                    Exit For
                End If
            Next
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in lbLookupResults_SelectedIndexChanged logging lookup")
            End If
        End If
        trPostCodeLookupOutput.Visible = False
        tbPostCodeLookup.Text = ""
        txtCneeName.Focus()
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration1CostCentre_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPerCustomerConfiguration1CostCentre.Text = ddlPerCustomerConfiguration1CostCentre.SelectedItem.Text
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration5CostCentre_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPerCustomerConfiguration5CostCentre.Text = ddlPerCustomerConfiguration5CostCentre.SelectedItem.Text
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration19CostCentre_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        tbPerCustomerConfiguration19CostCentre.Text = ddlPerCustomerConfiguration19CostCentre.SelectedItem.Text
    End Sub
    
    Protected Function sSetAuthorisationInfo(ByVal DataItem As Object) As String
        sSetAuthorisationInfo = String.Empty
        If pbOrderAuthorisation Or pbProductAuthorisation Then
            Dim au As AuthorisationInfo
            au = GetProductAuthorisationInfo(DataBinder.Eval(DataItem, "LogisticProductKey"))
            If au.bIsAuthorisable Then
                If CInt(au.sAvailableAuthorisation) > 0 Then
                    sSetAuthorisationInfo = au.sAvailableAuthorisation & " authorised"
                ElseIf au.bAuthorisationExpired Then
                    sSetAuthorisationInfo = "AUTHORISATION REQUIRED (previous authorisation has expired)"
                ElseIf au.sPendingAuthorisation <> "0" Then
                    sSetAuthorisationInfo = "AUTHORISATION REQUIRED (you have an authorisation request pending)"
                Else
                    sSetAuthorisationInfo = "AUTHORISATION REQUIRED"
                End If
            End If
        End If
        If pbCalendarManagement Then
            If Not IsDBNull(DataBinder.Eval(DataItem, "CalendarManaged")) Then
                Dim bCalendarManaged As Boolean = DataBinder.Eval(DataItem, "CalendarManaged")
                If bCalendarManaged Then
                    sSetAuthorisationInfo = "CALENDAR MANAGED"
                End If
            End If
        End If
        If pbCustomLetters Then
            If Not IsDBNull(DataBinder.Eval(DataItem, "CustomLetter")) Then
                Dim bCustomLetter As Boolean = DataBinder.Eval(DataItem, "CustomLetter")
                If bCustomLetter Then
                    sSetAuthorisationInfo = "CUSTOM LETTER"
                End If
            End If
        End If
    End Function
    
    Protected Function bSetAuthorisationInfoVisibility(ByVal DataItem As Object) As String
        bSetAuthorisationInfoVisibility = False
        If (pbOrderAuthorisation Or pbProductAuthorisation) Or pbCalendarManagement Or pbOnDemandProducts Then
            If pbOrderAuthorisation Or pbProductAuthorisation Then
                Dim au As AuthorisationInfo
                au = GetProductAuthorisationInfo(DataBinder.Eval(DataItem, "LogisticProductKey"))
                If au.bIsAuthorisable Then
                    bSetAuthorisationInfoVisibility = True
                End If
            Else
                bSetAuthorisationInfoVisibility = False
            End If
            ' bSetAuthorisationInfoVisibility = True   ' this may be always being set for reasons of improved visual appearance
            If pbCalendarManagement Then
                If Not IsDBNull(DataBinder.Eval(DataItem, "CalendarManaged")) Then
                    Dim bCalendarManaged As Boolean = DataBinder.Eval(DataItem, "CalendarManaged")
                    If bCalendarManaged Then
                        bSetAuthorisationInfoVisibility = True
                    End If
                End If
            End If
            If pbCustomLetters Then
                If Not IsDBNull(DataBinder.Eval(DataItem, "CustomLetter")) Then
                    Dim bCustomLetter As Boolean = DataBinder.Eval(DataItem, "CustomLetter")
                    If bCustomLetter Then
                        bSetAuthorisationInfoVisibility = True
                    End If
                End If
            End If
            If pbOnDemandProducts Then
                If Not IsDBNull(DataBinder.Eval(DataItem, "OnDemand")) Then
                    Dim nOnDemand As Integer = DataBinder.Eval(DataItem, "OnDemand")
                    If nOnDemand > 0 Then
                        bSetAuthorisationInfoVisibility = True
                    End If
                End If
            End If
        End If
    End Function

    Protected Sub lnkbtnGetFromSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetFromSharedAddressBook()
    End Sub
    
    Protected Sub GetFromSharedAddressBook()
        pbUsingSharedAddressBook = True
        dgAddressBook.CurrentPageIndex = 0
        lblLegendAddressBookType.Text = "Shared Address Book"
        lnkbtnUsePersonalAddressBook.Visible = True
        lnkbtnUseSharedAddressbook.Visible = False
        BindAddressBook()
        ShowSearchAddressListPanel()
    End Sub
    
    Protected Sub lnkbtnUseSharedAddressbook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSearchCriteriaAddress.Text = ""
        Call GetFromSharedAddressBook()
    End Sub

    Protected Sub lnkbtnUsePersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSearchCriteriaAddress.Text = ""
        Call GetFromPersonalAddressBook()
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration7ServiceLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        psServiceLevel = ddlPerCustomerConfiguration7ServiceLevel.SelectedValue
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration17ServiceLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        psServiceLevel = ddlPerCustomerConfiguration17ServiceLevel.SelectedItem.Text
        lblPerCustomerConfiguration17ConfirmationServiceLevel.Text = ddlPerCustomerConfiguration17ServiceLevel.SelectedItem.Text
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration18ServiceLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        psServiceLevel = ddlPerCustomerConfiguration18ServiceLevel.SelectedValue
    End Sub
    
    Protected Sub lnkbtnSaveAddressInPersonalAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddNewAddress()
        Call AddToPersonalAddressBook()
        Call HideSaveAddressLinks()
        If plCneeAddressKey > 0 Then
            WebMsgBox.Show("Address added to Personal Address Book")
        Else
            WebMsgBox.Show("The system encountered a problem adding the address to the Address Book")
        End If
    End Sub

    Protected Sub lnkbtnSaveAddressInSharedAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddNewAddress()
        Call AddToPersonalAddressBook()
        Call AddToSharedAddressBook()
        Call HideSaveAddressLinks()
        If plCneeAddressKey > 0 Then
            WebMsgBox.Show("Address added to Shared & Personal Address Book")
        Else
            WebMsgBox.Show("The system encountered a problem adding the address to the Address Book")
        End If
    End Sub

    Protected Sub HideSaveAddressLinks()
        lnkbtnSaveAddressInSharedAddressBook.Visible = False
        lnkbtnSaveAddressInPersonalAddressBook.Visible = False
    End Sub
    
    Protected Sub ShowSaveAddressLinks()
        lnkbtnSaveAddressInSharedAddressBook.Visible = True
        lnkbtnSaveAddressInPersonalAddressBook.Visible = True
    End Sub
    
    Protected Function gvProductListSetSubCategory(ByVal DataItem As Object) As String
        Dim sSubCategory As String = String.Empty
        Dim sSubSubCategory As String = String.Empty
        Dim sConjunction As String = String.Empty
        If Not IsDBNull(DataBinder.Eval(DataItem, "SubCategory")) Then
            sSubCategory = DataBinder.Eval(DataItem, "SubCategory")
        End If
        If Not IsDBNull(DataBinder.Eval(DataItem, "SubCategory2")) Then
            sSubSubCategory = DataBinder.Eval(DataItem, "SubCategory2")
        End If
        If sSubSubCategory <> String.Empty Then
            sConjunction = " \ "
        End If
        gvProductListSetSubCategory = sSubCategory & sConjunction & sSubSubCategory
    End Function
        
    Protected Function nGetAddressCount() As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim dtRecordCount As New DataTable
        Dim nRecordCount As Integer
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddressCount", oConn)
        Dim sSearchCriteria As String
        sSearchCriteria = txtSearchCriteriaAddress.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            If pbUsingSharedAddressBook Then
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

    Protected Function ReadAddressPage() As DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetAddressesPaged", oConn)
        Dim sSearchCriteria As String
        sSearchCriteria = txtSearchCriteriaAddress.Text
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            
            Dim nPageStart As Integer = ((pnAddressPage) * 20) + 1
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@PageStart", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@PageStart").Value = nPageStart
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@RowsToReturn", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@RowsToReturn").Value = 20
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
            If pbUsingSharedAddressBook Then
                oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
            Else
                oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")
            End If
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
            
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SortKey", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@SortKey").Value = "Country"

            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ""
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        ReadAddressPage = oDataTable
    End Function

    Protected Sub InitAddressDataGrid()
        pnAddressPage = START_ADDRESS_PAGE
        dgAddressBook.CurrentPageIndex = 0
        pnAddressVirtualItemCount = nGetAddressCount()
    End Sub
    
    Protected Function gvDistributionListSetQty(ByVal DataItem As Object, ByVal nColumn As Integer) As String
        Dim nBasketRow As Integer
        gvDistributionListSetQty = ""
        If nColumn <= gdtBasket.Rows.Count Then
            nBasketRow = nColumn - 1
            Dim dr As DataRow = gdtBasket.Rows(nBasketRow)
            gvDistributionListSetQty = dr("QtyToPick")
        End If
    End Function
    
    Protected Function sGetUserCostCentre() As String
        sGetUserCostCentre = ""
        Dim sSQL As String = "SELECT Department FROM UserProfile WHERE [key] = " & Session("UserKey")
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        Dim oDataTable As New DataTable
        oAdapter.Fill(oDataTable)
        sGetUserCostCentre = oDataTable.Rows(0).Item(0)
        oConn.Close()
        oDataTable.Dispose()
        oAdapter.Dispose()
    End Function

    Protected Function gvDistributionListSetCostCentre(ByVal DataItem As Object) As String
        gvDistributionListSetCostCentre = sGetUserCostCentre()
    End Function
    
    Protected Function GetDistributionListNames() As List(Of String)
        Dim lstDistributionListNames As New List(Of String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT DISTINCT DistributionListName FROM AddressDistributionLists WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY DistributionListName"
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                lstDistributionListNames.Add(oDataReader(0))
            End While
        Catch ex As Exception
            WebMsgBox.Show("GetDistributionListNames: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        Return lstDistributionListNames
    End Function
    
    Protected Function GetDistributionList(ByVal sListName As String) As DataTable ' HAD NO oConn.Open until 14MAY08
        Dim oConn As New SqlConnection(gsConn)
        Dim dtDistributionList As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Address_GetDistributionList", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DistributionListName", SqlDbType.NVarChar, 50))
            oAdapter.SelectCommand.Parameters("@DistributionListName").Value = sListName

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

            oConn.Open()
            oAdapter.Fill(dtDistributionList)
        Catch ex As Exception
            WebMsgBox.Show("GetDistributionList: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        GetDistributionList = dtDistributionList
    End Function
    
    Protected Sub ddlDistributionList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlDistributionList.SelectedIndex > 0 Then
            Call BindDistributionList()
            Call ShowDistributionListPanel()
        End If
    End Sub
    
    Protected Sub BindDistributionList()
        Dim oDataTable As DataTable = GetDistributionList(ddlDistributionList.SelectedItem.Text)
        Dim i As Integer
        For i = 0 To MULTIPLE_ADDRESS_PRODUCT_COLUMNS - 1
            gvDistributionList.Columns(i).Visible = False
        Next
        Call GetBasketFromSession()
        For i = 0 To gdtBasket.Rows.Count - 1
            Dim dr As DataRow = gdtBasket.Rows(i)
            gvDistributionList.Columns(i).HeaderText = dr("ProductCode")
            gvDistributionList.Columns(i).Visible = True
            'Select Case i
            '    Case 1 : tbItem1.text = dr("QtyToPick")
            'End Select
            If i = MULTIPLE_ADDRESS_PRODUCT_COLUMNS - 1 Then
                Exit For
            End If
        Next
        gvDistributionList.DataSource = oDataTable
        gvDistributionList.DataBind()
    End Sub

    Protected Sub gvDistributionList_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim gvDistributionList As GridView = sender

        Call GetBasketFromSession()
        Dim nBasketCount As Integer = gdtBasket.Rows.Count
        Dim i As Integer
        Dim nRow As Integer = 0
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                For i = nBasketCount To MULTIPLE_ADDRESS_PRODUCT_COLUMNS - 1
                    Dim tc As TableCell = gvr.Cells(i)
                    Dim cc As ControlCollection = tc.Controls
                    Dim cControlToRemove As Control = Nothing
                    For Each c As Control In cc
                        If TypeOf c Is TextBox Then
                            cControlToRemove = c
                            Exit For
                        End If
                    Next
                    cc.Remove(cControlToRemove)
                Next
                Dim lb As LinkButton
                lb = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("lnkbtnAddInfo")
                lb.CommandArgument = nRow
                nRow += 1
            End If
        Next
    End Sub

    Protected Function sSetProductOwnerInfo(ByVal DataItem As Object) As String
        sSetProductOwnerInfo = String.Empty
        If IsNumeric(DataItem.ToString) AndAlso DataItem.ToString > 0 Then
            sSetProductOwnerInfo = sGetProductOwnerDetails(DataItem.ToString)
        End If
    End Function
    
    Protected Function sGetProductOwnerDetails(ByVal sUserKey As String) As String
        sGetProductOwnerDetails = String.Empty
        Dim sFullName As String = String.Empty
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        sSQL = "SELECT FirstName, LastName, Telephone, EmailAddr FROM UserProfile WHERE [Key] = " & sUserKey
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                Dim sTelephone As String = String.Empty
                Dim sEmailAddr As String = String.Empty
                Dim sContactDetails As String = String.Empty
                If Not IsDBNull(oDataReader("Telephone")) Then
                    sTelephone = oDataReader("Telephone")
                    sTelephone = sTelephone.Trim
                End If
                If Not IsDBNull(oDataReader("EmailAddr")) Then
                    sEmailAddr = oDataReader("EmailAddr")
                    sEmailAddr = sEmailAddr.Trim
                End If
                If sTelephone.Length > 0 Then
                    sContactDetails = "tel: <b>" & sTelephone & "</b>&nbsp;&nbsp;"
                End If
                If sEmailAddr.Length > 0 Then
                    sContactDetails += "email: <b>" & sEmailAddr & "</b>"
                End If
                sFullName = oDataReader("FirstName") & " " & oDataReader("LastName")
                sFullName += " <a onmouseover=""return escape('" & sContactDetails & "')"" style='color: gray; cursor: help'> <img src='images/information_11x11.gif' /></a>"
            End If
            sGetProductOwnerDetails = sFullName
        Catch ex As Exception
            WebMsgBox.Show("sGetProductOwnerDetails: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Function gvProductListSetZeroStockNotificationVisibility(ByVal DataItem As Object) As String
        gvProductListSetZeroStockNotificationVisibility = False
        If pbZeroStockNotifications Then
            If DataBinder.Eval(DataItem, "Quantity") = 0 Then
                gvProductListSetZeroStockNotificationVisibility = True
            End If
        End If
    End Function

    Protected Function gvProductListShowUsage(ByVal DataItem As Object) As String
        gvProductListShowUsage = False
        Try
            If DataBinder.Eval(DataItem, "CalendarManaged") = True Then
                gvProductListShowUsage = True
            End If
        Catch
        End Try
    End Function
    
    Protected Function gvProductListSetAddToBasketVisibility(ByVal DataItem As Object) As String
        gvProductListSetAddToBasketVisibility = True
        If DataBinder.Eval(DataItem, "Quantity") = 0 Then
            gvProductListSetAddToBasketVisibility = False
        End If
    End Function
    
    Protected Function gvProductListGetProductStatusMessage() As String
        gvProductListGetProductStatusMessage = gsProductStatusMessage
    End Function

    Protected Function gvProductListGetLegend(ByVal sBaseName As String) As String
        gvProductListGetLegend = sBaseName
    End Function

    Protected Sub btnMultipleAddressOrderUpdateOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bRowUpdated As Boolean = False
        Dim hf As HiddenField
        Dim gvr As GridViewRow = gvDistributionList.Rows(hidEditRow.Value)
        tbMultipleAddressOrderCustomerRef.Text = tbMultipleAddressOrderCustomerRef.Text.Trim
        tbMultipleAddressOrderSpecialInstructions.Text = tbMultipleAddressOrderSpecialInstructions.Text.Trim
        tbMultipleAddressOrderShippingInfo.Text = tbMultipleAddressOrderShippingInfo.Text.Trim
        If tbMultipleAddressOrderCustomerRef.Text.Length > 0 Or tbMultipleAddressOrderSpecialInstructions.Text.Length > 0 Or tbMultipleAddressOrderShippingInfo.Text.Length > 0 Then
            bRowUpdated = True
        End If
        If cbUseCustomerRefForAllDestinations.Checked Then
            Call SetCustomerReferenceAllDestinations()
        Else
            hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidCustomerReference")
            hf.Value = tbMultipleAddressOrderCustomerRef.Text
        End If
        If cbUseSpecialInstructionsForAllDestinations.Checked Then
            Call SetSpecialInstructionsAllDestinations()
        Else
            hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidSpecialInstructions")
            hf.Value = tbMultipleAddressOrderSpecialInstructions.Text
        End If
        If cbUseShippingInfoForAllDestinations.Checked Then
            Call SetPackingNoteAllDestinations()
        Else
            hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidPackingNote")
            hf.Value = tbMultipleAddressOrderShippingInfo.Text
        End If
        Call UpdateAllRowColors()
        pnlUpdateOrder.Visible = False
    End Sub
    
    Protected Sub SetCustomerReferenceAllDestinations()
        Dim hf As HiddenField
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidCustomerReference")
                hf.Value = tbMultipleAddressOrderCustomerRef.Text
            End If
        Next
    End Sub
    
    Protected Sub SetSpecialInstructionsAllDestinations()
        Dim hf As HiddenField
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidSpecialInstructions")
                hf.Value = tbMultipleAddressOrderSpecialInstructions.Text
            End If
        Next
    End Sub

    Protected Sub SetPackingNoteAllDestinations()
        Dim hf As HiddenField
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidPackingNote")
                hf.Value = tbMultipleAddressOrderShippingInfo.Text
            End If
        Next
    End Sub

    Protected Sub UpdateAllRowColors()
        Dim hf As HiddenField
        Dim img As Image
        Dim nRow As Integer = 0
        Dim bRowUpdated As Boolean
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                bRowUpdated = False
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidCustomerReference")
                If hf.Value.Length > 0 Then
                    bRowUpdated = True
                End If
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidSpecialInstructions")
                If hf.Value.Length > 0 Then
                    bRowUpdated = True
                End If
                hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidPackingNote")
                If hf.Value.Length > 0 Then
                    bRowUpdated = True
                End If

                If bRowUpdated Then
                    img = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("imgTransparentSphere")
                    img.Visible = False
                    img = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("imgRedSphere")
                    img.Visible = True
                Else
                    img = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("imgTransparentSphere")
                    img.Visible = True
                    img = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("imgRedSphere")
                    img.Visible = False
                End If
                If nRow Mod 2 = 0 Then
                    gvr.BackColor = White
                Else
                    gvr.BackColor = WhiteSmoke
                End If
                nRow += 1
            End If
        Next
    End Sub
    
    Protected Sub lnkbtnAddInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim hf As HiddenField
        Dim gvr As GridViewRow
        Call UpdateAllRowColors()
        gvr = gvDistributionList.Rows(lnkbtn.CommandArgument)
        hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidCustomerReference")
        tbMultipleAddressOrderCustomerRef.Text = hf.Value
        hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidSpecialInstructions")
        tbMultipleAddressOrderSpecialInstructions.Text = hf.Value
        hf = gvr.Cells(MULTIPLE_ADDRESS_ADDRESSEE_COLUMN - 1).FindControl("hidPackingNote")
        tbMultipleAddressOrderShippingInfo.Text = hf.Value
        gvr.BackColor = Drawing.Color.FromArgb(&HFFFF99)
        tbMultipleAddressOrderCustomerRef.Focus()
        hidEditRow.Value = lnkbtn.CommandArgument
        cbUseCustomerRefForAllDestinations.Checked = False
        cbUseSpecialInstructionsForAllDestinations.Checked = False
        cbUseShippingInfoForAllDestinations.Checked = False
        pnlUpdateOrder.Visible = True
    End Sub
    
    Protected Sub ddlSetAllService_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            Call SetAllService(ddl.SelectedIndex - 1)
            ddl.SelectedIndex = 0
        End If
    End Sub

    Protected Sub SetAllService(ByVal nSelectedIndex As Integer)
        Dim ddl As DropDownList
        For Each gvr As GridViewRow In gvDistributionList.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                ddl = gvr.Cells(MULTIPLE_ADDRESS_SERVICE_LEVEL_COLUMN - 1).FindControl("ddlServiceLevel")
                ddl.SelectedIndex = nSelectedIndex
            End If
        Next
    End Sub
    
    Protected Function gvProductListGetUnitValue(ByVal DataItem As Object) As String
        gvProductListGetUnitValue = "Quantity:"
    End Function

    Protected Function IsShowingRichView() As Boolean
        IsShowingRichView = psProductView = PRODUCT_VIEW_CLASSIC
    End Function
    
    Protected Function ToggleProductView(ByVal sProductView As String) As String
        If sProductView = PRODUCT_VIEW_CLASSIC Then
            ToggleProductView = PRODUCT_VIEW_RICH
        Else
            ToggleProductView = PRODUCT_VIEW_CLASSIC
        End If
    End Function
    
    Protected Sub CreateSprintConfigCookie(ByVal sProductView As String)
        Dim c As HttpCookie = New HttpCookie("SprintConfig")
        c.Values.Add("SB_ProductView", sProductView)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
        'Response.Flush()
    End Sub
    
    Protected Sub CreateSprintBasketCookie()
        Dim c As HttpCookie = New HttpCookie("SprintBasket")
        c.Values.Add("SB_Basket", String.Empty)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub
    
    Protected Sub UpdateSprintConfigCookieProductView(ByVal sProductView As String)
        Dim c As HttpCookie = New HttpCookie("SprintConfig")
        'Dim c As HttpCookie = Request.Cookies("SprintConfig")
        c.Values.Add("SB_ProductView", sProductView)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
        'Response.Flush()
    End Sub
    
    Protected Sub FillBasketFromCookie()
        Dim gnrdicBasketItems As New Dictionary(Of Integer, Integer)
        Dim sBasket() As String = (Request.Cookies("SprintBasket")("SB_Basket") & "").Split("|")
        For Each s As String In sBasket
            If s.Contains(",") Then
                Dim sItem() As String = s.Split(",")
                If IsNumeric(sItem(0)) AndAlso (IsNumeric(sItem(1)) Or sItem(1) = String.Empty) Then
                    If sItem(1) = String.Empty Then
                        sItem(1) = "1" ' handle classic view case where items in basket but no quantity yet assigned
                    End If
                    gnrdicBasketItems.Add(CInt(sItem(0)), CInt(sItem(1)))
                End If
            End If
        Next
        Dim nAddedCount As Integer = 0
        Dim nRemovedCount As Integer = 0
        Dim bCheckedCustomerKey As Boolean = False
        For Each kv As KeyValuePair(Of Integer, Integer) In gnrdicBasketItems

            If Not bCheckedCustomerKey Then
                If IsThisCustomer(kv) Then
                    bCheckedCustomerKey = True
                Else
                    Exit Sub
                End If
            End If

            If ValidateCookieBasketItem(kv) Then
                nAddedCount += 1
                Call AddItemToBasket(kv.Key.ToString, bIsFromCookieBasket:=True)
                gdvBasketView = New DataView(gdtBasket)
                gdvBasketView.RowFilter = "ProductKey='" & kv.Key.ToString & "'"
                If gdvBasketView.Count = 1 AndAlso kv.Value > 0 Then
                    gdvBasketView.Item(0).Item("QtyToPick") = kv.Value
                End If
            Else
                nRemovedCount += 1
            End If
        Next
        Dim sbConfirmation As New StringBuilder
        Dim sPlural As String = String.Empty
        
        If nRemovedCount = 0 Then
            If nAddedCount <> 1 Then
                sbConfirmation.Append(nAddedCount.ToString & " items were left in your basket when a previous session ended. These have been placed in your basket again. You may remove them if they are no longer required.")
            Else
                sbConfirmation.Append(nAddedCount.ToString & " item was left in your basket when a previous session ended. This has been placed in your basket again. You may remove it if it is no longer required.")
            End If
        Else
            If gnrdicBasketItems.Count <> 1 Then
                sbConfirmation.Append(nAddedCount.ToString & " items were left in your basket when a previous session ended. ")
                'sbConfirmation.Append(nAddedCount.ToString & " item" & sPlural & " added to your basket. ")
            Else
                sbConfirmation.Append(nAddedCount.ToString & " item was left in your basket when a previous session ended. ")
                'sbConfirmation.Append(nAddedCount.ToString & " item" & sPlural & " added to your basket. ")
            End If
            If nAddedCount <> 1 Then
                sPlural = " items have "
            Else
                sPlural = " item has "
            End If
            sbConfirmation.Append(nAddedCount.ToString & sPlural & " been added to your basket. ")
            If nRemovedCount > 0 Then
                If nRemovedCount <> 1 Then
                    sPlural = "s"
                Else
                    sPlural = String.Empty
                End If
                sbConfirmation.Append(nRemovedCount.ToString & "item" & sPlural & " could not be added. This may be because of insufficient stock availability or because the product has been archived or deleted.")
            End If
        End If
        WebMsgBox.Show("Some items were left in your basket when a previous session ended. These have been placed in your basket again. You may remove them if they are no longer required.")
        WebMsgBox.Show(sbConfirmation.ToString)
    End Sub
    
    Protected Function IsThisCustomer(ByVal kv As KeyValuePair(Of Integer, Integer)) As Boolean
        IsThisCustomer = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("SELECT CustomerKey FROM LogisticProduct WHERE LogisticProductKey = " & kv.Key.ToString, oConn)
        Try
            oConn.Open()
            Dim oDataReader As SqlDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                If oDataReader("CustomerKey") = Session("CustomerKey") Then
                    IsThisCustomer = True
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in IsThisCustomer: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Function ValidateCookieBasketItem(ByVal kv As KeyValuePair(Of Integer, Integer)) As Boolean
        ValidateCookieBasketItem = True
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetFromKeyForCookieBasket2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int)
        oParamProductKey.Value = CLng(kv.Key)
        Dim oParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParamUserKey.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            'If oDataReader("ArchiveFlag") = "Y" OrElse oDataReader("DeletedFlag") = "Y" OrElse oDataReader("Quantity") <= CInt(kv.Value) Then    '  CN changed 5OCT09
            If oDataReader("ArchiveFlag") = "Y" OrElse oDataReader("DeletedFlag") = "Y" OrElse oDataReader("Quantity") < CInt(kv.Value) Then
                ValidateCookieBasketItem = False
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Error in ValidateBasketCookieItem: " & ex.Message)
        Finally
            If Not oDataReader Is Nothing Then
                oDataReader.Close()
            End If
            oConn.Close()
        End Try
    End Function
    
    Protected Sub SaveBasketToSession()
        If gnMode = MODE_JUPITER_STOCK Then
            Session("SB_BasketData") = gdtBasket
        Else
            Session("SB_BasketDataJupiter") = gdtBasket
        End If
        If Not gnMode = MODE_JUPITER_POD Then
            Call UpdateSprintBasketCookie()
        End If
    End Sub
        
    Protected Sub UpdateSprintBasketCookie()
        Dim sbEncodedBasket As New StringBuilder
        For Each dr As DataRow In gdtBasket.Rows
            sbEncodedBasket.Append(dr("ProductKey"))
            sbEncodedBasket.Append(",")
            sbEncodedBasket.Append(dr("QtyToPick"))
            sbEncodedBasket.Append("|")
        Next
        If sbEncodedBasket.Length > 0 Then
            sbEncodedBasket.Remove(sbEncodedBasket.Length - 1, 1)
        End If
        Dim c As HttpCookie = New HttpCookie("SprintBasket")
        c.Values.Add("SB_Basket", sbEncodedBasket.ToString)
        c.Expires = DateTime.Now.AddDays(2)
        Response.Cookies.Add(c)
    End Sub
    
    Protected Sub ClearSprintBasketCookie()
        Dim c As HttpCookie = New HttpCookie("SprintBasket")
        c.Values.Add("SB_Basket", String.Empty)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub
    
    Function GetProductAuthorisationInfo(ByVal sProductKey As String) As AuthorisationInfo
        Dim nAuthoriser As Integer = 0
        Dim au As New AuthorisationInfo()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd1 As New SqlCommand("SELECT * FROM LogisticProductAuthorisable WHERE LogisticProductKey = " & sProductKey, oConn)
        Try
            oConn.Open()
            Dim oDataReader1 As SqlDataReader = oCmd1.ExecuteReader()
            au.sAvailableAuthorisation = "0"
            au.sPendingAuthorisation = "0"
            If oDataReader1.Read() Then
                au.bIsAuthorisable = True
                nAuthoriser = oDataReader1("Authoriser")
                oDataReader1.Close()
                Dim ocmd2 As New SqlCommand("SELECT * FROM LogisticProductAuthorisation WHERE LogisticProductKey = " & sProductKey & " AND UserProfileKey = " & Session("UserKey"), oConn)
                Dim oDataReader2 As SqlDataReader = ocmd2.ExecuteReader()
                If oDataReader2.Read Then
                    au.bAuthorisationRecordFound = True
                    au.sID = oDataReader2("id")
                    Dim bGranted As Boolean = False
                    If Not IsDBNull(oDataReader2("Granted")) Then
                        bGranted = CBool(oDataReader2("Granted"))
                    End If
                    If bGranted Then
                        If IsDBNull(oDataReader2("QuantityRemaining")) OrElse CInt(oDataReader2("QuantityRemaining")) = 0 Then
                            au.sAvailableAuthorisation = "0"
                        Else
                            au.sAvailableAuthorisation = oDataReader2("QuantityRemaining")
                        End If
                        If Not IsDBNull(oDataReader2("AuthorisationExpiryDateTime")) Then
                            If CDate(oDataReader2("AuthorisationExpiryDateTime")) < Now() Then
                                au.sAvailableAuthorisation = "0"
                                au.dtAuthorisationExpiryDateTime = CDate(oDataReader2("AuthorisationExpiryDateTime"))
                                au.bAuthorisationExpired = True
                            End If
                        End If
                    Else
                        au.sPendingAuthorisation = oDataReader2("AuthorisedQuantity")
                        au.sAvailableAuthorisation = "0"
                    End If
                    'If Not IsDBNull(oDataReader2("Authoriser")) Then
                    '    au.nAuthoriser = CInt(oDataReader2("Authoriser"))
                    'End If
                    au.nAuthoriser = nAuthoriser
                End If
            Else
                au.bIsAuthorisable = False
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetProductAuthorisationInfo: " & ex.Message)
        Finally
            oConn.Close() ' CN 20NOV08
        End Try
        GetProductAuthorisationInfo = au
    End Function
    
    Protected Sub InitProductAuthorisationRequiredTable()
        gdtProductAuthorisationRequired = New DataTable
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("ProductKey", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("ProductCode", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("ProductDate", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("ProductDescription", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("Quantity", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("Notes", GetType(String)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("RequestAuthorisation", GetType(Boolean)))
        gdtProductAuthorisationRequired.Columns.Add(New DataColumn("Authoriser", GetType(Integer)))
    End Sub
    
    Protected Sub InitProductAuthorisationUsageTable()
        If IsNothing(Session("SB_AuthorisationUsage")) Then
            gdtAuthorisationUsage = New DataTable()
            gdtAuthorisationUsage.Columns.Add(New DataColumn("id", GetType(Long)))
            gdtAuthorisationUsage.Columns.Add(New DataColumn("QuantityUsed", GetType(Long)))
            Session("SB_AuthorisationUsage") = gdtAuthorisationUsage
        End If
        gdtAuthorisationUsage = Session("SB_AuthorisationUsage")
        gdtAuthorisationUsage.Rows.Clear()
    End Sub
    
    Protected Sub PrepareAuthorisations()
        Dim oConn As New SqlConnection(gsConn)
        Dim sAuthorisationNote As String = String.Empty
        Dim enumAuthorisationStatus As enumAuthorisationStatus
        Call InitProductAuthorisationUsageTable()
        Call InitProductAuthorisationRequiredTable()
        
        If Not (pbProductAuthorisation Or pbOrderAuthorisation) Then
            Exit Sub
        End If
        
        If Not UserMustAuthorise() Then
            Exit Sub
        End If
        
        For Each drBasket As DataRow In gdtBasket.Rows
            Dim au As AuthorisationInfo
            au = GetProductAuthorisationInfo(drBasket.Item("ProductKey"))
            If au.bIsAuthorisable Then
                'enumAuthorisationStatus = on_line_picks_aspx.enumAuthorisationStatus.NOT_AUTHORISED
                enumAuthorisationStatus = enumAuthorisationStatus.NOT_AUTHORISED
                sAuthorisationNote = "Authorisation required"
                If au.bAuthorisationRecordFound Then
                    If CInt(au.sPendingAuthorisation) > 0 Then
                        sAuthorisationNote = "Awaiting response to pending authorisation request"
                    ElseIf au.sAvailableAuthorisation = "0" Then
                        If au.bAuthorisationExpired Then
                            sAuthorisationNote = "Existing authorisation expired"
                        Else
                            sAuthorisationNote = "You have already ordered the quantity authorised"
                        End If
                    ElseIf CInt(drBasket.Item("QtyToPick")) > CInt(au.sAvailableAuthorisation) Then
                        sAuthorisationNote = "Your order for " & CInt(drBasket.Item("QtyToPick")).ToString & " units exceeds your remaining authorised quantity of " & au.sAvailableAuthorisation
                    Else
                        enumAuthorisationStatus = enumAuthorisationStatus.AUTHORISED
                        Dim drAuthorisationUsage As DataRow = gdtAuthorisationUsage.NewRow
                        drAuthorisationUsage.Item("id") = au.sID
                        drAuthorisationUsage.Item("QuantityUsed") = drBasket.Item("QtyToPick")
                        gdtAuthorisationUsage.Rows.Add(drAuthorisationUsage)
                    End If
                End If
            Else
                enumAuthorisationStatus = enumAuthorisationStatus.NOT_AUTHORISABLE
            End If
            
            If enumAuthorisationStatus = enumAuthorisationStatus.NOT_AUTHORISED Then
                Dim drProductAuthorisation As DataRow = gdtProductAuthorisationRequired.NewRow
                drProductAuthorisation.Item("ProductKey") = drBasket.Item("ProductKey")
                drProductAuthorisation.Item("ProductCode") = drBasket.Item("ProductCode")
                drProductAuthorisation.Item("ProductDate") = drBasket.Item("ProductDate")
                drProductAuthorisation.Item("ProductDescription") = drBasket.Item("Description")
                drProductAuthorisation.Item("Quantity") = drBasket.Item("QtyToPick")
                drProductAuthorisation.Item("Notes") = sAuthorisationNote
                drProductAuthorisation.Item("RequestAuthorisation") = True
                drProductAuthorisation.Item("Authoriser") = au.nAuthoriser
                gdtProductAuthorisationRequired.Rows.Add(drProductAuthorisation)
                drBasket.Item("Authorised") = "N"
            End If
        Next
        If gdtProductAuthorisationRequired.Rows.Count > 0 AndAlso pbOrderAuthorisation Then
            pbAuthorisationRequired = True
            pnAuthoriser = GetProductAuthoriser(CInt(gdtProductAuthorisationRequired.Rows(0).Item("ProductKey")))
            'lblAuthorisationAdvisory01.Text = ConfigLib.GetConfigItem_OrderAuthorisationAdvisory
            'lblAuthorisationAdvisory01.Visible = True
            gdtProductAuthorisationRequired.Clear()
            Call SaveBasketToSession()
            Exit Sub
        Else
            pbAuthorisationRequired = False
        End If
        
        If gdtProductAuthorisationRequired.Rows.Count > 0 Then
            gvRequestAuthorisation.DataSource = gdtProductAuthorisationRequired
            gvRequestAuthorisation.DataBind()
        End If
        oConn.Close()
        If gdtAuthorisationUsage.Rows.Count > 0 Then
            Session("SB_AuthorisationUsage") = gdtAuthorisationUsage
        End If
    End Sub
    
    Protected Function GetProductAuthoriser(ByVal nProductKey As Integer) As Integer
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderGetAuthoriser", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamProductKey As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int)
        oParamProductKey.Value = nProductKey
        GetProductAuthoriser = 0
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                oDataReader.Read()
                GetProductAuthoriser = oDataReader.Item(0)
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve product authoriser")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Function
    
    Protected Sub btnRequestAuthorisation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If AuthorisationRequestIsValid() Then
            Dim sMessage As String = String.Empty
            Call GetBasketFromSession()
            Dim nBasketPreAdjustmentItemCount As Integer = gdtBasket.Rows.Count
            Call ProcessAuthorisationsAndAdjustBasket()
            Call SaveBasketToSession()
            If gdtBasket.Rows.Count <> nBasketPreAdjustmentItemCount Then
                sMessage = "An authorisation request has been sent. Items requiring authorisation have been removed from your basket."
            End If
            If gdtBasket.Rows.Count > 0 Then
                Call CheckOrder()
            Else
                Call ShowProductList()
                'lblProductList.Text = "Product List"
                sMessage += " Your basket is now empty."
            End If
            If sMessage <> String.Empty Then
                Call WebMsgBox.Show(sMessage)
            End If
        Else
            Call WebMsgBox.Show("Validation failed.  Please check quantity fields are 0 or a positive number.")
        End If
    End Sub
    
    Protected Function AuthorisationRequestIsValid() As Boolean
        AuthorisationRequestIsValid = True
        For Each gvr As GridViewRow In gvRequestAuthorisation.Rows
            Dim tbAuthorisationQuantity As Object = Nothing
            tbAuthorisationQuantity = DirectCast(gvr.FindControl("tbAuthorisationQuantity"), TextBox)
            If Not IsNumeric(tbAuthorisationQuantity.Text) AndAlso CInt(tbAuthorisationQuantity.Text) > 0 Then
                AuthorisationRequestIsValid = False
            End If
        Next
    End Function
    
    Protected Sub ProcessAuthorisationsAndAdjustBasket()
        Dim oConn As New SqlConnection(gsConn)
        Dim hidLogisticProductKey As HiddenField
        Dim cbRequestAuthorisation As CheckBox
        Dim tbAuthorisationQuantity As TextBox
        Dim sLogisticProductKey As String = String.Empty
        Dim guidAuthorisationGUID As Guid
        
        For Each gvr As GridViewRow In gvRequestAuthorisation.Rows
            cbRequestAuthorisation = gvr.FindControl("cbRequestAuthorisation")
            hidLogisticProductKey = gvr.FindControl("hidLogisticProductKey")
            tbAuthorisationQuantity = gvr.FindControl("tbAuthorisationQuantity")
            If cbRequestAuthorisation.Checked Then
                sLogisticProductKey = hidLogisticProductKey.Value
                Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GenerateAuthorisationRequest", oConn)
                oCmd.CommandType = CommandType.StoredProcedure

                Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
                paramUserProfileKey.Value = Session("UserKey")
                oCmd.Parameters.Add(paramUserProfileKey)
                
                Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                paramLogisticProductKey.Value = sLogisticProductKey
                oCmd.Parameters.Add(paramLogisticProductKey)
                
                Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                paramCustomerKey.Value = Session("CustomerKey")
                oCmd.Parameters.Add(paramCustomerKey)

                guidAuthorisationGUID = Guid.NewGuid
                Dim paramAuthorisationGUID As SqlParameter = New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20)
                paramAuthorisationGUID.Value = guidAuthorisationGUID.ToString
                oCmd.Parameters.Add(paramAuthorisationGUID)

                Dim paramQuantity As SqlParameter = New SqlParameter("@Quantity", SqlDbType.Int, 4)
                paramQuantity.Value = tbAuthorisationQuantity.Text
                oCmd.Parameters.Add(paramQuantity)

                Dim sMessage As String = tbNoteToAuthoriser.Text.Trim
                Dim paramMessage As SqlParameter = New SqlParameter("@Message", SqlDbType.VarChar, 4000)
                If sMessage <> String.Empty Then
                    paramMessage.Value = sMessage
                Else
                    paramMessage.Value = System.Data.SqlTypes.SqlString.Null
                End If
                oCmd.Parameters.Add(paramMessage)

                Try
                    oConn.Open()
                    oCmd.Connection = oConn
                    oCmd.ExecuteNonQuery()

                Catch ex As SqlException
                    WebMsgBox.Show("Unable to send authorisation request - aborting")
                Finally
                    oConn.Close()
                End Try
                
                For Each dr As DataRow In gdtBasket.Rows
                    If dr.Item("ProductKey") = sLogisticProductKey Then
                        dr.Delete()
                        Exit For
                    End If
                Next
                gvBasket.DataBind()
                Session(gsBasketCountName) = gvBasket.Items.Count
                SetBasketCount(Session(gsBasketCountName))
            End If
        Next
    End Sub
    
    Protected Sub PlaceOrderOnHold()
        Dim lHoldingQueueKey As Long
        Dim guidAuthorisationGUID As Guid
        Dim drv As DataRowView
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderPlaceOnHold", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure

        guidAuthorisationGUID = Guid.NewGuid
        Dim paramAuthorisationGUID As SqlParameter = New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20)
        paramAuthorisationGUID.Value = guidAuthorisationGUID.ToString
        oCmdAddBooking.Parameters.Add(paramAuthorisationGUID)

        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CLng(Session("UserKey"))
        oCmdAddBooking.Parameters.Add(param1)

        Dim paramAuthoriserKey As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int, 4)
        paramAuthoriserKey.Value = pnAuthoriser
        oCmdAddBooking.Parameters.Add(paramAuthoriserKey)

        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CLng(Session("CustomerKey"))
        oCmdAddBooking.Parameters.Add(param2)

        Dim oBookingRefs As BookingRefs = GetBookingRefs()

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = oBookingRefs._BookingRef1
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = oBookingRefs._BookingRef2
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = oBookingRefs._BookingRef3
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = oBookingRefs._BookingRef4
        oCmdAddBooking.Parameters.Add(param6)
        
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param6a)

        Dim sSpecialInstructions As String = Session("SB_SpecialInstructions").ToString.Replace(vbCrLf, " ")
        
        If IsVSOE() AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery01.Checked Then
            sSpecialInstructions = "[AUTHORISED FOR EXPRESS DELIVERY] " & sSpecialInstructions
        End If
        
        If IsAAT() Then
            sSpecialInstructions = "[PRIORITY: " & ddlPerCustomerConfiguration17ServiceLevel.SelectedItem.Text & "] " & sSpecialInstructions
        End If

        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = sSpecialInstructions
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = Session("SB_ShippingNote")
        oCmdAddBooking.Parameters.Add(param8)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = psCnorCompany
        oCmdAddBooking.Parameters.Add(param13)

        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = psCnorAddr1
        oCmdAddBooking.Parameters.Add(param14)

        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = psCnorAddr2
        oCmdAddBooking.Parameters.Add(param15)

        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = psCnorAddr3
        oCmdAddBooking.Parameters.Add(param16)

        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = psCnorTown
        oCmdAddBooking.Parameters.Add(param17)

        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = psCnorState
        oCmdAddBooking.Parameters.Add(param18)

        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = psCnorPostCode
        oCmdAddBooking.Parameters.Add(param19)

        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = CLng(psCnorCountryKey)
        oCmdAddBooking.Parameters.Add(param20)

        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = psCnorCtcName
        oCmdAddBooking.Parameters.Add(param21)

        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = psCnorCtcTel
        oCmdAddBooking.Parameters.Add(param22)

        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = psCnorCtcEmail
        oCmdAddBooking.Parameters.Add(param23)

        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = Session("SB_CneeCompany")
        oCmdAddBooking.Parameters.Add(param25)

        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = Session("SB_CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)

        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = Session("SB_CneeAddr2")
        oCmdAddBooking.Parameters.Add(param27)

        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = Session("SB_CneeAddr3")
        oCmdAddBooking.Parameters.Add(param28)

        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = Session("SB_CneeTown")
        oCmdAddBooking.Parameters.Add(param29)

        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = Session("SB_CneeState")
        oCmdAddBooking.Parameters.Add(param30)

        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = Session("SB_CneePostCode")
        oCmdAddBooking.Parameters.Add(param31)

        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = Session("SB_CneeCountryKey")
        oCmdAddBooking.Parameters.Add(param32)

        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = Session("SB_CneeCtcName")
        oCmdAddBooking.Parameters.Add(param33)

        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = Session("SB_CneeCtcTel")
        oCmdAddBooking.Parameters.Add(param34)

        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = Session("SB_CneeCtcEmail")
        oCmdAddBooking.Parameters.Add(param35)

        Dim sAuthoriserMessageSingleAddressOrder As String = tbAuthoriserMessageSingleAddressOrder.Text.Replace(Environment.NewLine, " ")
        If IsVSOE() AndAlso rblPerCustomerConfiguration0CheckoutNextDayDelivery01.Checked Then
            sAuthoriserMessageSingleAddressOrder = "[SYSTEM MESSAGE: ORDER QUEUED FOR AUTHORISATION OF EXPRESS DELIVERY REQUEST. Order may also contain authorisable items.] " & sAuthoriserMessageSingleAddressOrder
        End If
        
        Dim param36 As SqlParameter = New SqlParameter("@MsgToAuthoriser", SqlDbType.NVarChar, 1000)
        param36.Value = sAuthoriserMessageSingleAddressOrder
        oCmdAddBooking.Parameters.Add(param36)

        Dim param37 As SqlParameter = New SqlParameter("@HoldingQueueKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)

        Try
            oConn.Open()
            oCmdAddBooking.ExecuteNonQuery()
            lHoldingQueueKey = CLng(oCmdAddBooking.Parameters("@HoldingQueueKey").Value.ToString)
            If lHoldingQueueKey > 0 Then
                Dim BasketView As New DataView
                Call GetBasketFromSession()
                BasketView = gdtBasket.DefaultView
                If BasketView.Count > 0 Then
                    For Each drv In BasketView
                        Try
                            Dim lProductKey As Long = CLng(drv("ProductKey"))
                            Dim lPickQuantity As Long = CLng(drv("QtyToPick"))
                            Dim sAuthorised As String
                            If Not IsDBNull(drv("Authorised")) Then
                                sAuthorised = drv("Authorised")
                            Else
                                sAuthorised = String.Empty
                            End If
                            Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderHoldingQueueItemAdd", oConn)
                            oCmdAddStockItem.CommandType = CommandType.StoredProcedure

                            Dim param53 As SqlParameter = New SqlParameter("@OrderHoldingQueueKey", SqlDbType.Int, 4)
                            param53.Value = lHoldingQueueKey
                            oCmdAddStockItem.Parameters.Add(param53)

                            Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                            param54.Value = lProductKey
                            oCmdAddStockItem.Parameters.Add(param54)

                            Dim param55 As SqlParameter = New SqlParameter("@Authorised", SqlDbType.Char, 1)
                            param55.Value = sAuthorised
                            oCmdAddStockItem.Parameters.Add(param55)

                            Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                            param56.Value = lPickQuantity
                            oCmdAddStockItem.Parameters.Add(param56)

                            oCmdAddStockItem.Connection = oConn
                            oCmdAddStockItem.ExecuteNonQuery()
                        Catch ex As Exception
                            WebMsgBox.Show(ex.ToString)
                        End Try
                    Next
                    If EmailAuthoriser(guidAuthorisationGUID.ToString) Then
                        Call ShowBookingQueuedConfirmationPanel()
                    End If
                Else
                    WebMsgBox.Show("Internal error - no product selected")
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show(ex.ToString)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function EmailAuthoriser(ByVal guidAuthorisationGUID As String) As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderGenerateRequest", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        
        EmailAuthoriser = True
        
        Dim paramUserProfileKey As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        paramUserProfileKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserProfileKey)
                
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramAuthoriserKey As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int, 4)
        paramAuthoriserKey.Value = pnAuthoriser
        oCmd.Parameters.Add(paramAuthoriserKey)

        guidAuthorisationGUID = guidAuthorisationGUID
        Dim paramAuthorisationGUID As SqlParameter = New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20)
        paramAuthorisationGUID.Value = guidAuthorisationGUID.ToString
        oCmd.Parameters.Add(paramAuthorisationGUID)

        Dim sMessage As String = tbAuthoriserMessageSingleAddressOrder.Text.Trim
        Dim paramMessage As SqlParameter = New SqlParameter("@Message", SqlDbType.VarChar, 1000)
        If sMessage <> String.Empty Then
            paramMessage.Value = sMessage
        Else
            paramMessage.Value = System.Data.SqlTypes.SqlString.Null
        End If
        oCmd.Parameters.Add(paramMessage)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Unable to send authorisation request - aborting")
            EmailAuthoriser = False
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub UpdateAuthorisations()
        If Not IsNothing(Session("SB_AuthorisationUsage")) Then
            gdtAuthorisationUsage = Session("SB_AuthorisationUsage")
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_AdjustAuthorisation", oConn)
            Dim spParam As SqlParameter
            oCmd.CommandType = CommandType.StoredProcedure
            For Each drAuthorisationUsage As DataRow In gdtAuthorisationUsage.Rows
                oCmd.Parameters.Clear()
                spParam = New SqlParameter("@AuthorisationKey", SqlDbType.Int, 4)
                spParam.Value = drAuthorisationUsage.Item("id")
                oCmd.Parameters.Add(spParam)
                spParam = New SqlParameter("@Increment", SqlDbType.Int, 4)
                spParam.Value = 0
                oCmd.Parameters.Add(spParam)
                spParam = New SqlParameter("@Decrement", SqlDbType.Int, 4)
                spParam.Value = drAuthorisationUsage.Item("QuantityUsed")
                oCmd.Parameters.Add(spParam)
                Try
                    oConn.Open()
                    oCmd.Connection = oConn
                    oCmd.ExecuteNonQuery()
                Catch ex As SqlException
                    WebMsgBox.Show("Unable to update authorisation")
                Finally
                    oConn.Close()
                End Try
            Next
        End If
    End Sub
    
    Protected Sub btnBackToDeliveryAddressFromSelectAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowDeliveryAddressPanel()
    End Sub
    
    Sub btn_ClassicRemoveBasketItems_click(ByVal sender As Object, ByVal e As EventArgs)
        Dim dgi As DataGridItem
        Dim cb As CheckBox
        lblError.Text = ""
        For Each dgi In dgrdBasket.Items
            cb = CType(dgi.Cells(1).Controls(1), CheckBox)
            If cb.Checked Then
                Call GetBasketFromSession()
                gdvBasketView = New DataView(gdtBasket)
                gdvBasketView.RowFilter = "ProductKey='" & dgi.Cells(0).Text & "'"
                If gdvBasketView.Count > 0 Then
                    gdvBasketView.Delete(0)
                End If
                gdvBasketView.RowFilter = ""
                Session(gsBasketCountName) = Session(gsBasketCountName) - 1
                SetBasketCount(Session(gsBasketCountName))
                Call SaveBasketToSession()
            End If
        Next dgi
        BindBasketGrid("ProductCode")
    End Sub
    
    Sub ClassicProductGrid_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "info" Then
            Dim itemCell As TableCell = e.Item.Cells(0)
            Dim lProductKey As Long = CLng(itemCell.Text)
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Product_GetProductFromKey", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As SqlParameter = oCmd.Parameters.Add("@ProductKey", SqlDbType.Int, 4)
            oParam.Value = lProductKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                If IsDBNull(oDataReader("ProductCode")) Then
                    lblProductCode.Text = ""
                Else
                    lblProductCode.Text = oDataReader("ProductCode")
                End If
                If IsDBNull(oDataReader("ProductDate")) Then
                    lblProductDate.Text = ""
                Else
                    lblProductDate.Text = oDataReader("ProductDate")
                End If
                If IsDBNull(oDataReader("ProductDescription")) Then
                    lblDescription.Text = ""
                Else
                    lblDescription.Text = oDataReader("ProductDescription")
                End If
                If IsDBNull(oDataReader("LanguageId")) Then
                    lblLanguage.Text = ""
                Else
                    lblLanguage.Text = oDataReader("LanguageId")
                End If
                If IsDBNull(oDataReader("ProductDepartmentId")) Then
                    lblDepartment.Text = ""
                Else
                    lblDepartment.Text = oDataReader("ProductDepartmentId")
                End If
                If IsDBNull(oDataReader("ProductCategory")) Then
                    lblCategory.Text = ""
                Else
                    lblCategory.Text = oDataReader("ProductCategory")
                End If
                If IsDBNull(oDataReader("SubCategory")) Then
                    lblSubCategory.Text = ""
                Else
                    lblSubCategory.Text = oDataReader("SubCategory")
                End If
                If IsDBNull(oDataReader("ItemsPerBox")) Then
                    lblItemsPerBox.Text = ""
                Else
                    lblItemsPerBox.Text = oDataReader("ItemsPerBox")
                End If
                If IsDBNull(oDataReader("MinimumStockLevel")) Then
                    lblMinStockLevel.Text = ""
                Else
                    lblMinStockLevel.Text = oDataReader("MinimumStockLevel")
                End If
                If IsDBNull(oDataReader("UnitValue")) Then
                    lblUnitValue.Text = ""
                Else
                    lblUnitValue.Text = Format(oDataReader("UnitValue"), "#,##0.00")
                End If
                If IsDBNull(oDataReader("UnitWeightGrams")) Then
                    lblUnitWeight.Text = ""
                Else
                    lblUnitWeight.Text = Format(oDataReader("UnitWeightGrams"), "#,##0")
                End If
                hlnk_DetailThumb.ImageUrl = psVirtualThumbFolder & oDataReader("ThumbNailImage")
                hlnk_DetailThumb.NavigateUrl = ConfigLib.GetConfigItem_Virtual_JPG_URL & oDataReader("OriginalImage")
                If oDataReader("PDFFileName") <> "blank_pdf.jpg" Then
                    hyplnkPDFDocument.Visible = True
                    hyplnkPDFDocument.NavigateUrl = ConfigLib.GetConfigItem_Virtual_PDF_URL & oDataReader("PDFFileName")
                Else
                    hyplnkPDFDocument.Visible = False
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
            Call ShowClassicProductDetail()
        End If
    End Sub
    
    Sub btn_ClassicReSelectCategory_click(ByVal s As Object, ByVal e As EventArgs)
        lblError.Text = String.Empty
        Call ShowCategoriesPanel()
    End Sub
    
    Protected Sub lnkbtnDisplayModeChange_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Session(gsBasketCountName) IsNot Nothing Then
            If Session(gsBasketCountName) > 0 Then
                WebMsgBox.Show("Cannot switch views while your basket contains products.")
                Exit Sub
            End If
        End If
        
        psProductView = ToggleProductView(psProductView)
        If IsShowingRichView() Then
            spanQuickModeCheckBox.Visible = True
        Else
            spanQuickModeCheckBox.Visible = False
        End If

        If pnlClassicProductList.Visible Then
            Call BindProductGridDispatcher("ProductCode")
            Call ShowProductList()
        End If

        If pnlClassicBasket.Visible Then
            Call BindBasketGrid("ProductCode")
            Call ShowBasket()
        End If

        If pnlProductList.Visible Then
            Call BindProductGridDispatcher("ProductCode")
            Call ShowClassicProductList()
        End If

        If pnlBasket.Visible Then
            Call BindBasketGrid("ProductCode")
            Call ShowClassicBasket()
        End If

    End Sub
    
    Protected Sub PerCustomerConfiguration0CheckoutNextDayDeliverySetCalendarVisibility()
        If rblPerCustomerConfiguration0CheckoutNextDayDelivery01.Checked Then
            spnPerCustomerConfiguration0CheckoutDeliveryDateCalendar.Visible = False
        Else
            spnPerCustomerConfiguration0CheckoutDeliveryDateCalendar.Visible = True
        End If
    End Sub

    Protected Sub rblPerCustomerConfiguration0CheckoutNextDayDelivery00_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PerCustomerConfiguration0CheckoutNextDayDeliverySetCalendarVisibility()
    End Sub

    Protected Sub rblPerCustomerConfiguration0CheckoutNextDayDelivery01_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PerCustomerConfiguration0CheckoutNextDayDeliverySetCalendarVisibility()
        If txtSpecialInstructions.Text.Contains("REQUIRED DELIVERY DATE") Then
            txtSpecialInstructions.Text = String.Empty
        End If
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration4Confirmation1ServiceLevel_Changed(ByVal sender As Object, ByVal e As EventArgs)
        If ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedIndex > 0 Then
            trPerCustomerConfiguration4Confirmation2.Visible = True
            psServiceLevel = ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedItem.Text
        Else
            trPerCustomerConfiguration4Confirmation2.Visible = False
        End If
        
        Call GetShippingCosts()
    End Sub

    Protected Sub ddlPerCustomerConfiguration25Confirmation1ServiceLevel_Changed(ByVal sender As Object, ByVal e As EventArgs)
        If ddlPerCustomerConfiguration25Confirmation1ServiceLevel.SelectedIndex > 0 Then
            trPerCustomerConfiguration25Confirmation2.Visible = True
            psServiceLevel = ddlPerCustomerConfiguration25Confirmation1ServiceLevel.SelectedItem.Text
            Call CABCostCalculator()
            Session("SB_BookingRef3") = "SVC: " & psServiceLevel
            Session("SB_BookingRef4") = "LOC: " & Session("UserName")
        Else
            trPerCustomerConfiguration25Confirmation2.Visible = False
        End If
    End Sub

    Protected Sub CABCostCalculator()
        Const COURIER_STANDARD As Double = 6.75
        Const COURIER_NEXTDAY As Double = 8.5
        Const COURIER_PER_KILO_ABOVE_10_KILOS As Double = 0.95
        'Const FUEL_SURCHARGE As Double = 1.08
        Const FUEL_SURCHARGE As Double = 1.0
        Const PACKAGING_ADDITION_GRAMS_PER_10_KILOS As Double = 500
        
        Dim dTotal As Double = 0.0
        Call GetBasketFromSession()
        Dim nTotalWeightGrams As Int32 = 0
        Dim nAdditionalKilos As Int32 = 0
        For Each dr As DataRow In gdtBasket.Rows
            nTotalWeightGrams += CInt(dr("UnitWeightGrams")) * CInt(dr("QtyToPick"))
        Next
        ' nTotalWeightGrams now holds total weight of ordered goods
        Dim nPackagingGramsToAdd As Double = PACKAGING_ADDITION_GRAMS_PER_10_KILOS
        Dim nTotalWeightGramsForPackaging As Int32 = nTotalWeightGrams
        While nTotalWeightGramsForPackaging > 10000
            nPackagingGramsToAdd += PACKAGING_ADDITION_GRAMS_PER_10_KILOS
            nTotalWeightGramsForPackaging -= 10000
        End While
        nTotalWeightGrams += nPackagingGramsToAdd      ' now have total weight INCLUDING packaging
        Dim nTotalWeightGramsForAdditionalKiloCalc As Int32 = nTotalWeightGrams
        While nTotalWeightGramsForAdditionalKiloCalc > 10000
            nAdditionalKilos += 1
            nTotalWeightGramsForAdditionalKiloCalc -= 1000
        End While
        ' nAdditionalKilos is integer number of kilos to be charged at additional rate
        If ddlPerCustomerConfiguration25Confirmation1ServiceLevel.SelectedValue = 1 Then
            dTotal = COURIER_STANDARD
        Else
            dTotal = COURIER_NEXTDAY
        End If
        dTotal = dTotal + (nAdditionalKilos * COURIER_PER_KILO_ABOVE_10_KILOS)
        dTotal = dTotal * FUEL_SURCHARGE
        lblPerCustomerConfiguration25Confirmation2ConsignmentWeight.Text = (nTotalWeightGrams / 1000).ToString
        lblLegendPerCustomerConfiguration25Confirmation2EstimagedShippingCost.Text = Format(dTotal, "#,##0.00")
    End Sub
    
    Public Class CostEstimate
        Public WeightCharge As Double
        Public EstimatedPackagingWeight As Double
        Public NonDoCSurCharge As Double
        Public DiscountRate As Double
        Public LocalTaxRate As Double
        Public Trace As String
    End Class

    Public Class CostCalculator
        
        Private sbCostEstimateTrace As New StringBuilder
        
        Public Function GetCostEstimate(ByVal lCustomerKey As Long, _
                                        ByVal lServiceLevelKey As Long, _
                                        ByVal sDocumentFlag As String, _
                                        ByVal sEstimatePackagingFlag As String, _
                                        ByVal lCountryKey As Long, _
                                        ByVal sTown As String, _
                                        ByVal sPostCode As String, _
                                        ByVal dblWeight As Double) As CostEstimate
    
            Dim dblWeightCharge As Double
            Dim dblMatrixBandFee As Double
            Dim bIsBaseRate As Boolean = True
            Dim dblBaseRate As Double
            Dim dblRemainder As Double
            Dim dblProductWeight As Double = dblWeight
            Dim dblPackagingWeight As Double = 0.0
    
            ' Create CustomerDetails Struct
            Dim oCostEstimate As CostEstimate = New CostEstimate()
            
            sbCostEstimateTrace.Append("TRACE @ " & DateTime.Now)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("CustKey=" & lCustomerKey)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("SvcLevKey=" & lServiceLevelKey)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("DocFlag=" & sDocumentFlag)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("EstPkg=" & sEstimatePackagingFlag)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("CtryKey=" & lCountryKey)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("Town=" & sTown)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("Postcode=" & sPostCode)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            sbCostEstimateTrace.Append("Weight=" & dblWeight)
            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
            
            Dim dr As DataRow
            Dim oConn As New SqlConnection(ConfigLib.GetConfigItem_ConnectionString())
            Dim oDataTable As New DataTable
            Dim oAdapter As New SqlDataAdapter("spASPNET_Tariff_GetZoneMatrixFromAddress", oConn)
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            Try
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = lCustomerKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ServiceLevelKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ServiceLevelKey").Value = lServiceLevelKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DocumentFlag", SqlDbType.NVarChar, 1))
                oAdapter.SelectCommand.Parameters("@DocumentFlag").Value = sDocumentFlag
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CountryKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CountryKey").Value = lCountryKey
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Town", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@Town").Value = sTown
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@PostalCode", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@PostalCode").Value = sPostCode
    
                oAdapter.Fill(oDataTable)


                If sEstimatePackagingFlag = "Y" Then
                    Do While dblProductWeight > 0
                        dblPackagingWeight = dblPackagingWeight + 1.25
                        dblProductWeight = dblProductWeight - 12.5
                    Loop
                    
                    sbCostEstimateTrace.Append("Estimated packaging weight: " & dblPackagingWeight.ToString)
                    sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                    
                    dblWeight = dblWeight + dblPackagingWeight
                End If
                
                Dim nIteration As Int32 = 0
                
                sbCostEstimateTrace.Append("TARIFF: WeightFrom ~ WeightTo ~ Fee ~ Units ~ FlatRate ~ HoldBase ~ NonDocSurcharge ~ DiscountRate ~ LocalTaxRate")
                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                For Each dr In oDataTable.Rows
                    nIteration += 1
                    sbCostEstimateTrace.Append(nIteration.ToString & ": " & dr("WeightFrom") & ", " & dr("WeightTo") & ", " & dr("Fee") & ", " & dr("Units") & ", " & dr("FlatRate") & ", " & dr("HoldBase") & ", " & dr("NonDocSurcharge") & ", " & dr("DiscountRate") & ", " & dr("LocalTaxRate"))
                    sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                Next

                sbCostEstimateTrace.Append(oDataTable.Rows.Count & " record(s) retrieved")
                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                
                nIteration = 0
                
                For Each dr In oDataTable.Rows
                    nIteration += 1
                    sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                    sbCostEstimateTrace.Append("ITERATION " & nIteration.ToString)
                    sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                    If bIsBaseRate Then
                        dblBaseRate = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")
                        bIsBaseRate = False
                        sbCostEstimateTrace.Append("Calculated base rate as " & dblBaseRate.ToString)
                        sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                    End If
                    
                    If dblWeight >= dr("WeightTo") Then
                        
                        sbCostEstimateTrace.Append("Weight (" & dblWeight.ToString & ") > 'WeightTo' (" & dr("WeightTo") & ")")
                        sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                        
                        If dr("FlatRate") = False Then  ' not a flat rate

                            sbCostEstimateTrace.Append("Not flat rate")
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            dblMatrixBandFee = ((dr("WeightTo") - dr("WeightFrom")) / dr("Units")) * dr("Fee")

                            sbCostEstimateTrace.Append("MatrixBandFee {(WeightTo - WeightFrom) / Units * Fee} = " & dblMatrixBandFee.ToString & ", Units = " & dr("Units") & ", Fee = " & dr("Fee"))
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            dblWeightCharge = dblWeightCharge + dblMatrixBandFee
                        Else                       ' this is now (possible already was) a flat rate charge

                            sbCostEstimateTrace.Append("Flat rate")
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            If dr("HoldBase") = True Then

                                sbCostEstimateTrace.Append("Hold base")
                                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else

                                sbCostEstimateTrace.Append("NOT Hold base")
                                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    ElseIf dblWeight >= dr("WeightFrom") And dblWeight < dr("WeightTo") Then

                        sbCostEstimateTrace.Append("Weight (" & dblWeight.ToString & ") between 'WeightFrom' (" & dr("WeightFrom") & ") and 'WeightTo' ( " & dr("WeightTo") & ")")
                        sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                        If dr("FlatRate") = False Then  ' not a flat rate

                            sbCostEstimateTrace.Append("Not flat rate")
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            dblRemainder = (dblWeight - dr("WeightFrom")) / dr("Units")

                            sbCostEstimateTrace.Append("Remainder {(Weight - WeightFrom) / Units)} = " & dblRemainder.ToString & ", Units = " & dr("Units"))
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            Do While dblRemainder > 0
                                dblWeightCharge = dblWeightCharge + dr("Fee")

                                sbCostEstimateTrace.Append("Decrementing remainder = " & dblRemainder.ToString)
                                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                                dblRemainder = dblRemainder - 1
                            Loop
                        Else

                            sbCostEstimateTrace.Append("Flat rate")
                            sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                            If dr("HoldBase") = True Then  'see above for explanation

                                sbCostEstimateTrace.Append("Hold base")
                                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                                dblWeightCharge = ((dblWeight / dr("Units")) * dr("Fee")) + dblBaseRate
                            Else

                                sbCostEstimateTrace.Append("NOT Hold base")
                                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                                dblWeightCharge = (dblWeight / dr("Units")) * dr("Fee")
                            End If
                        End If
                    End If
                    oCostEstimate.NonDoCSurCharge = CDbl(dr("NonDocSurcharge"))
                    oCostEstimate.DiscountRate = CDbl(dr("DiscountRate"))
                    oCostEstimate.LocalTaxRate = CDbl(dr("LocalTaxRate"))
                    
                    sbCostEstimateTrace.Append("END ITERATION " & nIteration & ", Weight = " & dblWeight.ToString & ", WeightCharge = " & dblWeightCharge.ToString & ", PackagingWeight = " & dblPackagingWeight)
                    sbCostEstimateTrace.Append("<br />" & Environment.NewLine)

                Next
    
                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                sbCostEstimateTrace.Append("END CALCULATION, Weight = " & dblWeight.ToString & ", PackagingWeight = " & dblPackagingWeight & ", Total Weight = " & (dblWeight + dblPackagingWeight).ToString & ", WeightCharge = " & dblWeightCharge.ToString & ", NonDoCSurCharge = " & oCostEstimate.NonDoCSurCharge & ", DiscountRate = " & oCostEstimate.DiscountRate & ", LocalTaxRate  = " & oCostEstimate.LocalTaxRate)
                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                
                sbCostEstimateTrace.Append("For zone selection calculation see sproc spASPNET_Tariff_GetZoneMatrixFromAddress; for possible tariffs for this country use SELECT * FROM TariffDestination WHERE CountryKey = ++country key++")
                sbCostEstimateTrace.Append("<br />" & Environment.NewLine)
                '

                oCostEstimate.WeightCharge = dblWeightCharge
                oCostEstimate.EstimatedPackagingWeight = dblPackagingWeight
                oCostEstimate.Trace = sbCostEstimateTrace.ToString
    
            Catch ex As SqlException
            Finally
                oConn.Close()
            End Try
            Return oCostEstimate
        End Function
    End Class
    
    Protected Sub GetShippingCosts()
        If ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedItem.Value = 0 Then
            Exit Sub
        End If
        Dim oCostCalculator As CostCalculator = New CostCalculator()
        Dim oCostEstimate As CostEstimate = New CostEstimate()
        
        Dim dblProductWeight As Double
        Dim dblEstimatedPackagingWeight As Double
        Dim dblWeightCharge As Double
        Dim dblDiscountRate As Double
        Dim dblDiscountAmount As Double
        Dim dblDiscountedCharge As Double
        Dim dblNDS As Double
        Dim dblSubTotal As Double
        Dim dblLocalTaxRate As Double
        Dim dblLocalTaxAmount As Double
        Dim dblTotal As Double
        Dim nTempCustomerKey As Integer = Session("CustomerKey")
        If nTempCustomerKey = CUSTOMER_YALE Then
            nTempCustomerKey = CUSTOMER_HYSTER
        End If
        
        dblProductWeight = (plBasketWeightGrams / 1000)
        
        Session("DelService") = ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedItem.Text
        ' note: 4th param was originally parameterised as ConfigurationManager.AppSettings("EstimatePackaging")
        oCostEstimate = oCostCalculator.GetCostEstimate(nTempCustomerKey, _
                                                            CLng(ddlPerCustomerConfiguration4Confirmation1ServiceLevel.SelectedItem.Value), _
                                                            "N", _
                                                            "Y", _
                                                            Session("SB_CneeCountryKey"), _
                                                            Session("SB_CneeTown"), _
                                                            Session("SB_CneePostCode"), _
                                                            dblProductWeight)
        
        dblEstimatedPackagingWeight = oCostEstimate.EstimatedPackagingWeight
        dblWeightCharge = oCostEstimate.WeightCharge
        dblDiscountRate = oCostEstimate.DiscountRate
        dblDiscountAmount = (oCostEstimate.WeightCharge * oCostEstimate.DiscountRate) / 100
        dblDiscountedCharge = dblWeightCharge + dblDiscountAmount
        dblNDS = oCostEstimate.NonDoCSurCharge
        dblSubTotal = (dblDiscountedCharge + dblNDS) * 1.08   ' 8% fuel surcharge - should really either read from table or put in web.config
        dblLocalTaxRate = Format(oCostEstimate.LocalTaxRate, "#,##0.00")
        dblLocalTaxAmount = (dblSubTotal * dblLocalTaxRate) / 100
        dblTotal = dblSubTotal + dblLocalTaxAmount
        
        lblPerCustomerConfiguration4Confirmation2Weight.Text = Format((dblProductWeight + dblEstimatedPackagingWeight), "#,##0.0")
        lblPerCustomerConfiguration4Confirmation2BasketShippingCost.Text = Format(dblSubTotal, "#,##0.00")
        
        Call ExecuteQueryToDataTable("INSERT INTO CostEstimateTrace (Result, DateTime, UserKey) VALUES ('" & oCostEstimate.Trace.Replace("'", "''") & "', GETDATE(), 0)")
    End Sub
    
    Protected Function CheckBasketForCustomLetter() As Integer   ' assume dtBasket is populated since only here if ValidBasket()
        CheckBasketForCustomLetter = 0
        Call CreateBasketIfNull()
        Call GetBasketFromSession()
        For Each dr As DataRow In gdtBasket.Rows
            If Not IsDBNull(dr("CustomLetter")) AndAlso dr("CustomLetter") Then  ' need to work out why first condition is necessary
                CheckBasketForCustomLetter = CInt(dr("ProductKey"))
                Exit For
            End If
        Next
    End Function

    Protected Function CheckBasketForCalendarManagedItems() As Boolean   ' assume dtBasket is populated since only here if ValidBasket()
        CheckBasketForCalendarManagedItems = False
        Call CreateBasketIfNull()
        Call GetBasketFromSession()
        For Each dr As DataRow In gdtBasket.Rows
            If Not IsDBNull(dr("CalendarManaged")) AndAlso dr("CalendarManaged") Then  ' need to work out why first condition is necessary
                CheckBasketForCalendarManagedItems = True
                Exit For
            End If
        Next
    End Function

    Protected Function GetCalendarManagedItems() As List(Of Integer)
        Dim nlstCalendarManagedItems As New List(Of Integer)
        Call GetBasketFromSession()
        For Each dr As DataRow In gdtBasket.Rows
            If Not IsDBNull(dr("CalendarManaged")) AndAlso dr("CalendarManaged") Then  ' need to work out why first condition is necessary
                nlstCalendarManagedItems.Add(dr("ProductKey"))
            End If
        Next
        GetCalendarManagedItems = nlstCalendarManagedItems
    End Function

    Protected Function GetCalendarManagedItemsDataView() As DataView
        Dim dvCalendarManagedItems As DataView
        Call GetBasketFromSession()
        dvCalendarManagedItems = New DataView(gdtBasket)
        dvCalendarManagedItems.RowFilter = "CalendarManaged='True'"
        Return dvCalendarManagedItems
    End Function
    
    Protected Sub ShowCalendarManagedPanel(ByVal bSkipClearDateSelection As Boolean)
        Dim dvCalendarManagedItems As DataView = GetCalendarManagedItemsDataView()
        Call HideAllPanels()
        pnlCalendarManaged.Visible = True
        'dtBasket = Session("SB_BasketData")
        'dvCalendarManagedItems = New DataView(dtBasket)
        'dvCalendarManagedItems.RowFilter = "CalendarManaged='True'"
        gvCalendarManagedItems.DataSource = dvCalendarManagedItems
        gvCalendarManagedItems.DataBind()
        If dvCalendarManagedItems.Count = 1 Then
            lblCMLegendCalendar.Text = "Select day(s) on which product is required"
            lblCMLegendViewOtherReservations.Text = "Events using this product"
            btnCMBookEvent.Text = "book product for event"
        Else
            lblCMLegendCalendar.Text = "Select day(s) on which products are required"
            lblCMLegendViewOtherReservations.Text = "Events using these products"
            btnCMBookEvent.Text = "book products for event"
        End If
        Call RetrieveBookings(New Date(DateTime.Now.Year, DateTime.Now.Month, 1)) ' need this call since pDictBookings will not have been set
        If Not bSkipClearDateSelection Then
            Call CMClearDateSelection()
        End If
        Call InitOtherCMEventGrid(dvCalendarManagedItems)
        calCalendar1.PrevMonthText = String.Empty
    End Sub

    Protected Sub RemoveCMItemsFromBasket()
        Const CM_MESSAGE As String = "Your booking has been received (you can view your bookings on the My Profile tab). Calendar Managed products have been removed from your basket. "
        Dim nlstCalendarManagedItems As List(Of Integer) = GetCalendarManagedItems()
        For Each nLogisticProductKey As Integer In nlstCalendarManagedItems
            Call RemoveItemFromBasket(nLogisticProductKey)
        Next
        If gdtBasket.Rows.Count = 0 Then
            Call WebMsgBox.Show(CM_MESSAGE & "Your basket is now empty.")
            Call ShowCategoriesPanel()
        Else
            If gdtBasket.Rows.Count = 1 Then
                Call WebMsgBox.Show(CM_MESSAGE & "Please continue your order with the remaining product.")
            Else
                Call WebMsgBox.Show(CM_MESSAGE & "Please continue your order with the remaining products.")
            End If
            Call PaintSessionVariables()
            Call CheckOrder()
        End If
    End Sub
    
    Protected Sub InitOtherCMEventGrid(ByVal dvCalendarManagedItems As DataView)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL1 As New StringBuilder
        Dim bFirstTime As Boolean = True
        sbSQL1.Append("SELECT DISTINCT EventId INTO #EventId FROM CalendarManagedItemDays cmid INNER JOIN CalendarManagedItemEvent cmie ON cmid.EventId = cmie.id WHERE EventDay >= GETDATE() AND ISNULL(IsDeleted,0) = 0 AND (")
        For Each drv As DataRowView In dvCalendarManagedItems
            If Not bFirstTime Then
                sbSQL1.Append(" OR ")
            End If
            sbSQL1.Append(" LogisticProductKey = ")
            sbSQL1.Append(drv("ProductKey"))
            bFirstTime = False
        Next
        sbSQL1.Append(")")
        Dim oCmd1 As SqlCommand = New SqlCommand(sbSQL1.ToString, oConn)
        Try
            oConn.Open()
            oCmd1.ExecuteNonQuery()
        
            Dim sbSQL2 As New StringBuilder
            sbSQL2.Append("SELECT EventName Event, FirstName + ' ' + LastName 'Booked by', CONVERT(VARCHAR(9), MIN(EventDay), 6) 'Delivery Date', CONVERT(VARCHAR(9), MAX(EventDay), 6) 'Collection Date' ")
            sbSQL2.Append("FROM #EventId ei INNER JOIN CalendarManagedItemDays cmid ON ei.EventId = cmid.EventId INNER JOIN CalendarManagedItemEvent cmie ON ei.EventId = cmie.id INNER JOIN UserProfile up ON cmie.BookedBy = up.[Key] ")
            sbSQL2.Append("GROUP BY EventName, FirstName, LastName ")
            sbSQL2.Append("ORDER BY MIN(EventDay) ")
            Dim oCmd2 As SqlCommand = New SqlCommand(sbSQL2.ToString, oConn)

            oDataReader = oCmd2.ExecuteReader()
            Dim arr As ArrayList = New ArrayList
            For Each row As Object In oDataReader
                arr.Add(row)
            Next

            gvOtherCalendarManagedReservations.DataSource = arr
            gvOtherCalendarManagedReservations.DataBind()
        Catch ex As Exception
            WebMsgBox.Show("InitOtherCMEventGrid: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function bCalendarManagedNoDateSelected() As Boolean
        Dim bResult As Boolean
        bResult = pdtCalendarManagedSelectionX = Nothing And pdtCalendarManagedSelectionY = Nothing
        If bResult Then
            btnCMBookEvent.Enabled = False
        Else
            btnCMBookEvent.Enabled = True
        End If
        bCalendarManagedNoDateSelected = bResult
    End Function
    
    Protected Sub calCalendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim c As System.Web.UI.WebControls.Calendar = sender
        Dim dtSelectedDate As Date = c.SelectedDate
        Dim dtBookingWarningDate As Date = DateAdd(DateInterval.Day, 2, Date.Today)
        Dim bIsAccountHandler As Boolean = CBool(Session("UserPermissions") And USER_PERMISSION_ACCOUNT_HANDLER)
        If dtSelectedDate <= Date.Today Then
            If Not (bIsAccountHandler And (dtSelectedDate = Date.Today)) Then
                WebMsgBox.Show("Only future bookings are accepted")
                Exit Sub
            End If
        End If
        If Not bIsAccountHandler Then
            If dtSelectedDate <= dtBookingWarningDate Then
                WebMsgBox.Show("Online bookings for delivery next day can normally be accepted up to 3.00pm the previous working day. Please confirm short notice bookings with Customer Services.")
            End If
        End If
        If bCalendarManagedNoDateSelected() Then
            gdictCurrentMonthBookings = pdictCurrentMonthBookings
            'Call RetrieveBookings(calCalendar1.VisibleDate)  ' could optimise this to check only for this date
            If gdictCurrentMonthBookings.ContainsKey(dtSelectedDate) Then
                WebMsgBox.Show("One or more of the items you have selected is already in use on this date")
            Else
                pdtCalendarManagedSelectionX = dtSelectedDate
                pdtCalendarManagedSelectionY = dtSelectedDate
                Call MarkSelectedDays()
                lnkbtnFindAvailableProducts.Visible = True
            End If
        Else
            If dtSelectedDate < pdtCalendarManagedSelectionX Then
                If Not BookingsInDateRange(dtSelectedDate, pdtCalendarManagedSelectionY) Then
                    pdtCalendarManagedSelectionX = dtSelectedDate
                Else
                    WebMsgBox.Show("One or more of the items you have selected is already in use during this period")
                End If
                Call MarkSelectedDays()
            Else
                If Not BookingsInDateRange(pdtCalendarManagedSelectionX, dtSelectedDate) Then
                    pdtCalendarManagedSelectionY = dtSelectedDate
                Else
                    WebMsgBox.Show("One or more of the items you have selected is already in use during this period")
                End If
                Call MarkSelectedDays()
            End If
        End If
    End Sub

    Protected Sub MarkSelectedDays()
        If Not bCalendarManagedNoDateSelected() Then
            Dim d As New Date
            d = pdtCalendarManagedSelectionX
            glstSelection = New List(Of Date)
            Do
                glstSelection.Add(d)
                d = DateAdd(DateInterval.Day, 1, d)
            Loop Until d > pdtCalendarManagedSelectionY
        End If
    End Sub

    Protected Function BookingsInDateRange(ByVal dtStartDate As Date, ByVal dtEndDate As Date) As Boolean
        If IsNothing(gdictSelectionBookings) Then
            gdictSelectionBookings = New Dictionary(Of Date, String)
            Dim nlstCalendarManagedItems As List(Of Integer) = GetCalendarManagedItems()
            For Each nLogisticProductKey As Integer In nlstCalendarManagedItems
                Call RetrieveBookingForProduct(gdictSelectionBookings, nLogisticProductKey, dtStartDate, dtEndDate)
            Next
        End If
        BookingsInDateRange = gdictSelectionBookings.Count > 0
    End Function
    
    Protected Sub calCalendar1_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs)
        Dim dtBookingWarningDate As Date = DateAdd(DateInterval.Day, 2, Date.Today)
        Dim c As System.Web.UI.WebControls.Calendar = sender
        Dim drea As DayRenderEventArgs = e
        Dim cd As CalendarDay = drea.Day
        Dim tc As TableCell = drea.Cell
        If gdictCurrentMonthBookings Is Nothing Then
            Dim dtBaseDate As Date = DateAdd(DateInterval.Day, 10, e.Day.Date)
            Call RetrieveBookings(New Date(dtBaseDate.Year, dtBaseDate.Month, 1))
        End If
        If cd.Date < Date.Today Then
            tc.BackColor = Gray
        End If
        If (cd.Date >= Date.Today) And (cd.Date < dtBookingWarningDate) Then
            tc.BackColor = LightGray
        End If
        If cd.Date < dtBookingWarningDate Then
            tc.Enabled = False
        End If
        
        If gdictCurrentMonthBookings.ContainsKey(cd.Date) Then
            tc.BackColor = Red
        End If
        If glstSelection IsNot Nothing Then
            If glstSelection.Contains(cd.Date) Then
                tc.BackColor = Green
            End If
        End If
    End Sub
    
    Protected Sub calCalendar1_VisibleMonthChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs)
        Dim c As System.Web.UI.WebControls.Calendar = sender
        Dim mcea As MonthChangedEventArgs = e
        Call RetrieveBookings(mcea.NewDate)
        Call MarkSelectedDays()
        If (mcea.NewDate.Month > Date.Today.Month) Or (mcea.NewDate.Year > Date.Today.Year) Then
            calCalendar1.PrevMonthText = "&lt;"
        Else
            calCalendar1.PrevMonthText = String.Empty
        End If
    End Sub
    
    Protected Sub RetrieveBookings(ByVal dtBaseDate As Date)   ' need to ensure base date always arrives as first of month
        Dim dictBooking As New Dictionary(Of Date, String)
        Dim dtStartDate As Date = DateAdd(DateInterval.Day, -6, dtBaseDate)
        Dim dtEndDate As Date = DateAdd(DateInterval.Month, 1, dtBaseDate)
        dtEndDate = DateAdd(DateInterval.Day, 14, dtEndDate) ' calendar shows 6 weeks; worst case is 28 day feb, showing 2 weeks of following month
        ' for now do each product separately, but later optimise database call by doing them together, passing list of products as CSV to SQL Server
        Dim nlstCalendarManagedItems As List(Of Integer) = GetCalendarManagedItems()
        For Each nLogisticProductKey As Integer In nlstCalendarManagedItems
            Call RetrieveBookingForProduct(dictBooking, nLogisticProductKey, dtStartDate, dtEndDate)
        Next
        gdictCurrentMonthBookings = dictBooking
        Dim d = New Dictionary(Of Date, String)
        For Each kv As KeyValuePair(Of Date, String) In dictBooking
            d.Add(kv.Key, kv.Value)
        Next
        pdictCurrentMonthBookings = d
    End Sub
    
    Protected Sub RetrieveBookingForProduct(ByRef dictBooking As Dictionary(Of Date, String), ByVal nLogisticProductKey As Integer, ByVal dtStartDate As Date, ByVal dtEndDate As Date)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_GetEventsInDateRange2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = nLogisticProductKey
        
        Dim paramStartDate As SqlParameter = oCmd.Parameters.Add("@StartDate", SqlDbType.SmallDateTime)
        paramStartDate.Value = dtStartDate

        Dim paramEndDate As SqlParameter = oCmd.Parameters.Add("@EndDate", SqlDbType.SmallDateTime)
        paramEndDate.Value = dtEndDate

        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read()
                    Try
                        dictBooking.Add(oDataReader("EventDay"), oDataReader("LogisticProductKey") & "," & oDataReader("EventName"))
                    Catch ' duplicate keys
                    End Try
                End While
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve event information (spASPNET_CalendarManaged_GetEventsInDateRange)")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnCMClearDateSelection_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CMClearDateSelection()
        lnkbtnFindAvailableProducts.Visible = False
    End Sub
    
    Protected Sub CMClearDateSelection()
        pdtCalendarManagedSelectionX = Nothing
        pdtCalendarManagedSelectionY = Nothing
        calCalendar1.SelectedDate = Nothing
        Dim bTemp As Boolean = bCalendarManagedNoDateSelected() ' hide button
    End Sub

    Protected Sub ClearCMFields()
        Call CMClearDateSelection()
        tbCMContactName.Text = String.Empty
        tbCMContactPhone.Text = String.Empty
        tbCMContactMobile.Text = String.Empty
        tbCMEventName.Text = String.Empty
        tbCMEventAddress1.Text = String.Empty
        tbCMEventAddress2.Text = String.Empty
        tbCMTown.Text = String.Empty
        tbCMPostcode.Text = String.Empty
        cbCMDifferentCollectionAddress.Checked = False
        tbCMCollectionAddress1.Text = String.Empty
        tbCMCollectionAddress2.Text = String.Empty
        tbCMCollectionTown.Text = String.Empty
        tbCMCollectionPostcode.Text = String.Empty
        ddlCMItemDeliverBy.SelectedIndex = 0
        tbCMExactDeliveryPoint.Text = String.Empty
        ddlCMItemCollectBetween.SelectedIndex = 0
        tbCMExactCollectionPoint.Text = String.Empty
        tbCMSpecialInstructions.Text = String.Empty
        tbCMCustomerReference.Text = String.Empty
        Call SetCollectionFieldsVisibility(False)
    End Sub
    
    Protected Sub btnCMBookEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TryBookEvent()
    End Sub
    
    Protected Function bIsUniqueEventNameForCustomer() As Boolean
        bIsUniqueEventNameForCustomer = False
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_GetEventNameForCustomer", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        
        Dim oParamCustomerKey As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int)
        oParamCustomerKey.Value = Session("CustomerKey")
        
        Dim oEventName As SqlParameter = oCmd.Parameters.Add("@EventName", SqlDbType.VarChar, 50)
        oEventName.Value = tbCMEventName.Text
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If Not oDataReader.HasRows Then
                bIsUniqueEventNameForCustomer = True
            Else
                WebMsgBox.Show("The event name you entered has already been used - please choose another name")
            End If
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve event name (spASPNET_CalendarManaged_GetEventNameForCustomer)")
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    End Function
    
    Protected Sub TryBookEvent()
        Validate("CalendarManaged")
        If Not (Page.IsValid AndAlso bIsUniqueEventNameForCustomer()) Then
            Call MarkSelectedDays()
            Exit Sub
        End If
        If Not BasketContainsBookableItems() Then
            WebMsgBox.Show("Your basket contains no bookable calendar-managed items.")
            Exit Sub
        End If
        Call RemoveNonBookableItemsFromBasket()
        Call CMBookEvent()
        ' Call CMClearDateSelection()
        Call MarkSelectedDays()
        Call InitOtherCMEventGrid(GetCalendarManagedItemsDataView())
        If Not cbCMMultipleBookings.Checked Then
            Call RemoveCMItemsFromBasket()
        End If
    End Sub
    
    Protected Sub RemoveNonBookableItemsFromBasket()
        Call GetBasketFromSession()
        For i As Integer = (gdtBasket.Rows.Count - 1) To 0 Step -1
            Dim dr As DataRow = gdtBasket.Rows(i)
            If CBool(dr("CalendarManaged")) Then
                If dr("Notes").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN1) Or dr("ProductDate").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN2) Then
                    gdtBasket.Rows(i).Delete()
                    Session(gsBasketCountName) = Session(gsBasketCountName) - 1
                End If
            End If
        Next
        Call SaveBasketToSession()
    End Sub
    
    Protected Function BasketContainsBookableItems() As Boolean
        BasketContainsBookableItems = False
        Call GetBasketFromSession()
        For Each dr As DataRow In gdtBasket.Rows
            If CBool(dr("CalendarManaged")) Then
                If Not (dr("Notes").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN1) Or dr("ProductDate").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN2)) Then
                    BasketContainsBookableItems = True
                    Exit For
                End If
            End If
        Next
    End Function
    
    Protected Sub CMBookEvent()
        If pdtCalendarManagedSelectionY < Date.Today Then
            WebMsgBox.Show("WARNING: The collection date selected for this event has already passed!")
        Else
            If pdtCalendarManagedSelectionX < Date.Today Then
                WebMsgBox.Show("WARNING: The delivery date selected for this event has already passed!")
            End If
        End If
        If glstSelection Is Nothing Then
            Call MarkSelectedDays()
        End If
        Dim nEventId As Integer
        nEventId = CreateCMEvent()
        Dim nlstCalendarManagedItems As List(Of Integer) = GetCalendarManagedItems()
        For Each nLogisticProductKey As Integer In nlstCalendarManagedItems
            For Each d As Date In glstSelection
                Call AddItemDateToEvent(nLogisticProductKey, nEventId, d)
            Next
        Next
        Dim bOverseasBooking As Boolean = False
        If ddlCMCountry.Visible Then
            If ddlCMCountry.SelectedValue <> COUNTRY_CODE_UK Then
                bOverseasBooking = True
            End If
        End If
        If ddlCMCollectionCountry.Visible Then
            If ddlCMCollectionCountry.SelectedValue <> COUNTRY_CODE_UK Then
                bOverseasBooking = True
            End If
        End If
        If bOverseasBooking Then
            Call AlertOverseasBooking()
        End If
        Call ClearCMFields()
    End Sub
    
    Protected Sub AlertOverseasBooking()
        Dim sSQL As String
        Dim sRecipientName As String
        Dim sRecipientEmail As String
        Dim sText As String
        sSQL = "SELECT ISNULL(ah.Name,'Someone'), ISNULL(ah.EmailAddr,'') FROM AccountHandler ah INNER JOIN Customer c ON ah.[key] = c.AccountHandlerKey WHERE c.CustomerKey = " & Session("CustomerKey") & " AND ah.DeletedFlag <> 1"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            sRecipientName = dr(0)
            sRecipientEmail = dr(1).ToString.Trim
            If sRecipientEmail = String.Empty Then
                sRecipientEmail = "account.managers@transworld.eu.com"
            End If
        Else
            sRecipientName = "Account Handler"
            sRecipientEmail = "account.managers@transworld.eu.com"
        End If
        sText = "User " & Session("UserName") & ", Customer " & Session("Customer") & ", has just booked an Event outside the UK. The event name is " & tbCMEventName.Text
        Call SendMail("OVERSEAS EVENT ALERT", sRecipientEmail, "OVERSEAS EVENT SYSTEM ALERT", sText, sText)
    End Sub
    
    Protected Sub AddItemDateToEvent(ByVal nLogisticProductKey As Integer, ByVal nEventId As Integer, ByVal d As Date)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_CMAddItemDateToEventBooking", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = nLogisticProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        
        Dim paramEventDay As SqlParameter = New SqlParameter("@EventDay", SqlDbType.SmallDateTime)
        paramEventDay.Value = d
        oCmd.Parameters.Add(paramEventDay)
        
        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = nEventId
        oCmd.Parameters.Add(paramEventId)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show("Could not add item to event due to internal error (spASPNET_Product_CMAddItemDateToEventBooking)" & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function CreateCMEvent() As Integer
        Dim guidAccessGUID As Guid = Guid.NewGuid
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_CMAddEventBooking5", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.NVarChar, 50)
        paramContactName.Value = tbCMContactName.Text
        oCmd.Parameters.Add(paramContactName)

        Dim paramContactPhone As SqlParameter = New SqlParameter("@ContactPhone", SqlDbType.VarChar, 50)
        paramContactPhone.Value = tbCMContactPhone.Text
        oCmd.Parameters.Add(paramContactPhone)

        Dim paramContactMobile As SqlParameter = New SqlParameter("@ContactMobile", SqlDbType.VarChar, 50)
        paramContactMobile.Value = tbCMContactMobile.Text
        oCmd.Parameters.Add(paramContactMobile)

        Dim paramContactName2 As SqlParameter = New SqlParameter("@ContactName2", SqlDbType.NVarChar, 50)
        paramContactName2.Value = tbCMContactName2.Text
        oCmd.Parameters.Add(paramContactName2)

        Dim paramContactPhone2 As SqlParameter = New SqlParameter("@ContactPhone2", SqlDbType.VarChar, 50)
        paramContactPhone2.Value = tbCMContactPhone2.Text
        oCmd.Parameters.Add(paramContactPhone2)

        Dim paramContactMobile2 As SqlParameter = New SqlParameter("@ContactMobile2", SqlDbType.VarChar, 50)
        paramContactMobile2.Value = tbCMContactMobile2.Text
        oCmd.Parameters.Add(paramContactMobile2)

        Dim paramEventName As SqlParameter = New SqlParameter("@EventName", SqlDbType.NVarChar, 50)
        paramEventName.Value = tbCMEventName.Text
        oCmd.Parameters.Add(paramEventName)

        Dim paramEventAddress1 As SqlParameter = New SqlParameter("@EventAddress1", SqlDbType.NVarChar, 50)
        paramEventAddress1.Value = tbCMEventAddress1.Text
        oCmd.Parameters.Add(paramEventAddress1)

        Dim paramEventAddress2 As SqlParameter = New SqlParameter("@EventAddress2", SqlDbType.NVarChar, 50)
        paramEventAddress2.Value = tbCMEventAddress2.Text
        oCmd.Parameters.Add(paramEventAddress2)

        Dim paramEventAddress3 As SqlParameter = New SqlParameter("@EventAddress3", SqlDbType.NVarChar, 50)
        paramEventAddress3.Value = String.Empty
        oCmd.Parameters.Add(paramEventAddress3)

        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.NVarChar, 50)
        paramTown.Value = tbCMTown.Text
        oCmd.Parameters.Add(paramTown)

        Dim paramPostcode As SqlParameter = New SqlParameter("@Postcode", SqlDbType.NVarChar, 50)
        paramPostcode.Value = tbCMPostcode.Text
        oCmd.Parameters.Add(paramPostcode)

        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int)
        If ddlCMCountry.Visible Then
            paramCountryKey.Value = ddlCMCountry.SelectedValue
        Else
            paramCountryKey.Value = COUNTRY_CODE_UK
        End If
        oCmd.Parameters.Add(paramCountryKey)
        
        Dim paramDeliveryTime As SqlParameter = New SqlParameter("@DeliveryTime", SqlDbType.VarChar, 50)
        paramDeliveryTime.Value = ddlCMItemDeliverBy.SelectedValue
        oCmd.Parameters.Add(paramDeliveryTime)

        Dim paramPreciseDeliveryPoint As SqlParameter = New SqlParameter("@PreciseDeliveryPoint", SqlDbType.NVarChar, 100)
        paramPreciseDeliveryPoint.Value = tbCMExactDeliveryPoint.Text
        oCmd.Parameters.Add(paramPreciseDeliveryPoint)

        Dim paramDifferentCollectionAddress As SqlParameter = New SqlParameter("@DifferentCollectionAddress", SqlDbType.Bit)
        paramDifferentCollectionAddress.Value = cbCMDifferentCollectionAddress.Checked
        oCmd.Parameters.Add(paramDifferentCollectionAddress)

        Dim paramCollectionAddress1 As SqlParameter = New SqlParameter("@CollectionAddress1", SqlDbType.NVarChar, 50)
        paramCollectionAddress1.Value = tbCMCollectionAddress1.Text
        oCmd.Parameters.Add(paramCollectionAddress1)

        Dim paramCollectionAddress2 As SqlParameter = New SqlParameter("@CollectionAddress2", SqlDbType.NVarChar, 50)
        paramCollectionAddress2.Value = tbCMCollectionAddress2.Text
        oCmd.Parameters.Add(paramCollectionAddress2)

        Dim paramCollectionTown As SqlParameter = New SqlParameter("@CollectionTown", SqlDbType.NVarChar, 50)
        paramCollectionTown.Value = tbCMCollectionTown.Text
        oCmd.Parameters.Add(paramCollectionTown)

        Dim paramCollectionPostcode As SqlParameter = New SqlParameter("@CollectionPostcode", SqlDbType.NVarChar, 50)
        paramCollectionPostcode.Value = tbCMCollectionPostcode.Text
        oCmd.Parameters.Add(paramCollectionPostcode)

        Dim paramCollectionCountryKey As SqlParameter = New SqlParameter("@CollectionCountryKey", SqlDbType.Int)
        If ddlCMCollectionCountry.Visible Then
            paramCollectionCountryKey.Value = ddlCMCollectionCountry.SelectedValue
        Else
            paramCollectionCountryKey.Value = COUNTRY_CODE_UK
        End If
        oCmd.Parameters.Add(paramCollectionCountryKey)
        
        Dim paramCollectionTime As SqlParameter = New SqlParameter("@CollectionTime", SqlDbType.VarChar, 50)
        paramCollectionTime.Value = ddlCMItemCollectBetween.SelectedValue
        oCmd.Parameters.Add(paramCollectionTime)

        Dim paramPreciseCollectionPoint As SqlParameter = New SqlParameter("@PreciseCollectionPoint", SqlDbType.NVarChar, 100)
        paramPreciseCollectionPoint.Value = tbCMExactCollectionPoint.Text
        oCmd.Parameters.Add(paramPreciseCollectionPoint)

        Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 200)
        paramSpecialInstructions.Value = tbCMSpecialInstructions.Text
        oCmd.Parameters.Add(paramSpecialInstructions)

        Dim paramCustomerReference As SqlParameter = New SqlParameter("@CustomerReference", SqlDbType.NVarChar, 100)
        paramCustomerReference.Value = tbCMCustomerReference.Text
        oCmd.Parameters.Add(paramCustomerReference)

        Dim paramAccessGUID As SqlParameter = New SqlParameter("@AccessGUID", SqlDbType.VarChar, 30)
        paramAccessGUID.Value = guidAccessGUID.ToString
        oCmd.Parameters.Add(paramAccessGUID)

        Dim paramBookedBy As SqlParameter = New SqlParameter("@BookedBy", SqlDbType.Int)
        paramBookedBy.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramBookedBy)

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Direction = ParameterDirection.Output
        oCmd.Parameters.Add(paramEventId)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
            CreateCMEvent = CLng(oCmd.Parameters("@EventId").Value)
        Catch ex As SqlException
            WebMsgBox.Show("Could not create event due to internal error: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub dgrdProducts_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dg As DataGrid = sender
        Dim dgiea As DataGridItemEventArgs = e
        Dim dgi As DataGridItem = dgiea.Item
        'If dgi.ItemType = ListItemType.Header Then
        '            If IsHyster() Then
        ' dgi.Cells(7).Text = "Value (€)"
        'Else
        '    dgi.Cells(7).Text = "Value (£)"
        'End If
        'End If
        If dgi.ItemType = ListItemType.Item Or dgi.ItemType = ListItemType.AlternatingItem Then
            If IsHysterOrYale() Then
                dgi.Cells(7).Text = "€" & dgi.Cells(7).Text
            Else
                dgi.Cells(7).Text = "£" & dgi.Cells(7).Text
            End If
        End If
        If (Session("UserPermissions") And USER_PERMISSION_VIEW_STOCK) Then
            dgi.Cells(11).Visible = False
        End If

    End Sub

    Protected Sub dgrdBasket_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dg As DataGrid = sender
        Dim dgiea As DataGridItemEventArgs = e
        Dim dgi As DataGridItem = dgiea.Item
        If dgi.ItemType = ListItemType.Item Or dgi.ItemType = ListItemType.AlternatingItem Then
            If IsHysterOrYale() Then
                dgi.Cells(6).Text = "€" & dgi.Cells(6).Text
            Else
                dgi.Cells(6).Text = "£" & dgi.Cells(6).Text
            End If
        End If

    End Sub

    Protected Function sGetSessionGUID() As String
        sGetSessionGUID = psOnDemandSessionGUID
    End Function
        
    Protected Sub btnZeroStockNotification_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hidProductkey As HiddenField = sender.NamingContainer.FindControl("hidProductkey")
        Dim sProductKey As String = hidProductkey.Value.ToString
        Dim sEmailAddress As String = GetUserEmailAddress()
        Call RecordNotificationAddressAndUser(sProductKey, sEmailAddress, 0)
        WebMsgBox.Show("A notification email will be sent to your registered email address (currently " & sEmailAddress & ") when this product becomes available.\n\nYou can view the notifications you have requested on the My Profile tab.")
    End Sub

    Protected Function GetUserEmailAddress() As String
        GetUserEmailAddress = String.Empty
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetProfileFromKey3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If oDataReader.HasRows Then
                If Not IsDBNull(oDataReader("EmailAddr")) Then
                    GetUserEmailAddress = oDataReader("EmailAddr").ToString.Trim
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in GetUserEmailAddress: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub RecordNotificationAddressAndUser(ByVal nLogisticProductKey As Integer, ByVal sEmailAddr As String, ByVal nQuantityRequired As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_ZeroStockNotification_Record", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramLogisticProductKey As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int)
        paramLogisticProductKey.Value = nLogisticProductKey
        oCmd.Parameters.Add(paramLogisticProductKey)
        
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int)
        paramUserKey.Value = Session("UserKey")
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
    
    Protected Sub lnkbtnCustomLetterResetFromTemplate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PreConfirmOrder()
    End Sub

    Protected Sub cbCMDifferentCollectionAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCollectionFieldsVisibility(cb.Checked)
        If Not cb.Checked Then
            tbCMCollectionAddress1.Text = String.Empty
            tbCMCollectionAddress2.Text = String.Empty
            tbCMCollectionTown.Text = String.Empty
            tbCMCollectionPostcode.Text = String.Empty
            rfvCMCollectionAddress1.Enabled = False
            rfvCMCollectionTown.Enabled = False
            rfvCMCollectionPostcode.Enabled = False
        Else
            rfvCMCollectionAddress1.Enabled = True
            rfvCMCollectionTown.Enabled = True
            rfvCMCollectionPostcode.Enabled = True
        End If
    End Sub
    
    Protected Sub SetCollectionFieldsVisibility(ByVal bVisibility As Boolean)
        trCMCollectionAddress1.Visible = bVisibility
        trCMCollectionAddress2.Visible = bVisibility
        trCMCollectionTown.Visible = bVisibility
        trCMCollectionPostcode.Visible = bVisibility
        'trCMCollectionCountry.Visible = bVisibility
    End Sub
    
    Protected Sub lnkbtnCMFindEventAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CMFindEventAddress()
    End Sub

    Protected Sub CMFindEventAddress()
        tbCMPostcode.Text = tbCMPostcode.Text.Trim.ToUpper
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult
        objInterimResults = objLookup.ByPostcode(tbCMPostcode.Text, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        If objInterimResults.IsError OrElse objInterimResults.Results Is Nothing OrElse objInterimResults.Results.GetLength(0) = 0 Then
            lblCMFindEventAddressFailure.Visible = True
            trCMSelectEventAddress.Visible = False
            WebMsgBox.Show(objInterimResults.ErrorMessage)
        Else
            lblCMFindEventAddressFailure.Visible = False
            trCMSelectEventAddress.Visible = True
            lbCMSelectEventAddress.Items.Clear()
            If Not objInterimResults.Results Is Nothing Then
                For Each objInterimResult In objInterimResults.Results
                    lbCMSelectEventAddress.Items.Add(New ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in CMFindEventAddress logging lookup")
            End If
            trCMSelectEventAddress.Visible = True
        End If
    End Sub
    
    Protected Sub lnkbtnCMFindCollectionAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CMFindCollectionAddress()
    End Sub
    
    Protected Sub CMFindCollectionAddress()
        tbCMCollectionPostcode.Text = tbCMCollectionPostcode.Text.Trim.ToUpper
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objInterimResults As uk.co.postcodeanywhere.services.InterimResults
        Dim objInterimResult As uk.co.postcodeanywhere.services.InterimResult
        objInterimResults = objLookup.ByPostcode(tbCMCollectionPostcode.Text, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        If objInterimResults.IsError OrElse objInterimResults.Results Is Nothing OrElse objInterimResults.Results.GetLength(0) = 0 Then
            lblCMFindCollectionAddressFailure.Visible = True
            trCMSelectCollectionAddress.Visible = False
            WebMsgBox.Show(objInterimResults.ErrorMessage)
        Else
            lblCMFindCollectionAddressFailure.Visible = False
            trCMSelectCollectionAddress.Visible = True
            lbCMSelectCollectionAddress.Items.Clear()
            If Not objInterimResults.Results Is Nothing Then
                For Each objInterimResult In objInterimResults.Results
                    lbCMSelectCollectionAddress.Items.Add(New ListItem(objInterimResult.Description, objInterimResult.Id))
                Next
            End If
            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in CMFindEventAddress logging lookup")
            End If
            trCMSelectCollectionAddress.Visible = True
        End If
    End Sub
    
    Protected Sub lnkbtnCMCancelSelectEventAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMSelectEventAddress.Visible = False
    End Sub

    Protected Sub lnkbtnCMCancelSelectCollectionAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMSelectCollectionAddress.Visible = False
    End Sub

    Protected Sub lbCMSelectEventAddress_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CMSelectEventAddress()
    End Sub
    
    Protected Sub CMSelectEventAddress()
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objAddressResults As uk.co.postcodeanywhere.services.AddressResults
        Dim objAddress As uk.co.postcodeanywhere.services.Address

        objAddressResults = objLookup.FetchAddress(lbCMSelectEventAddress.SelectedValue, uk.co.postcodeanywhere.services.enLanguage.enLanguageEnglish, uk.co.postcodeanywhere.services.enContentType.enContentStandardAddress, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        If objAddressResults.IsError Then
            WebMsgBox.Show(objAddressResults.ErrorMessage)
        Else
            objAddress = objAddressResults.Results(0)
            'txtCneeName.Text = objAddress.OrganisationName
            tbCMEventAddress1.Text = objAddress.Line1
            tbCMEventAddress2.Text = objAddress.Line2 & " " & objAddress.Line3
            tbCMTown.Text = objAddress.PostTown
            tbCMPostcode.Text = objAddress.Postcode
            'txtCneeState.Text = objAddress.County

            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in lbLookupResults_SelectedIndexChanged logging lookup")
            End If
        End If
        trCMSelectEventAddress.Visible = False
    End Sub

    Protected Sub lbCMSelectCollectionAddress_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CMSelectCollectionAddress()
    End Sub
    
    Protected Sub CMSelectCollectionAddress()
        Dim objLookup As New uk.co.postcodeanywhere.services.LookupUK
        Dim objAddressResults As uk.co.postcodeanywhere.services.AddressResults
        Dim objAddress As uk.co.postcodeanywhere.services.Address

        objAddressResults = objLookup.FetchAddress(lbCMSelectCollectionAddress.SelectedValue, uk.co.postcodeanywhere.services.enLanguage.enLanguageEnglish, uk.co.postcodeanywhere.services.enContentType.enContentStandardAddress, ACCOUNT_CODE, LICENSE_KEY, "")
        objLookup.Dispose()
        If objAddressResults.IsError Then
            WebMsgBox.Show(objAddressResults.ErrorMessage)
        Else
            objAddress = objAddressResults.Results(0)
            'txtCneeName.Text = objAddress.OrganisationName
            tbCMCollectionAddress1.Text = objAddress.Line1
            tbCMCollectionAddress2.Text = objAddress.Line2 & " " & objAddress.Line3
            tbCMCollectionTown.Text = objAddress.PostTown
            tbCMCollectionPostcode.Text = objAddress.Postcode
            'txtCneeState.Text = objAddress.County

            Dim sHostIPAddress As String = Server.HtmlEncode(Request.UserHostName)
            If Not ExecuteNonQuery("INSERT INTO PostCodeLookup (ClientIPAddress, LookupDateTime, CustomerKey, CostInUnits) VALUES ('" & sHostIPAddress & "', GETDATE(), " & Session("CustomerKey") & ", 1)") Then
                WebMsgBox.Show("Error in lbLookupResults_SelectedIndexChanged logging lookup")
            End If
        End If
        trCMSelectCollectionAddress.Visible = False
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

    Property psCnorCompany() As String
        Get
            Dim o As Object = ViewState("SB_CnorCompany")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCompany") = Value
        End Set
    End Property
    
    Property psCnorAddr1() As String
        Get
            Dim o As Object = ViewState("SB_CnorAddr1")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorAddr1") = Value
        End Set
    End Property
    
    Property psCnorAddr2() As String
        Get
            Dim o As Object = ViewState("SB_CnorAddr2")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorAddr2") = Value
        End Set
    End Property
    
    Property psCnorAddr3() As String
        Get
            Dim o As Object = ViewState("SB_CnorAddr3")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorAddr3") = Value
        End Set
    End Property
    
    Property psCnorTown() As String
        Get
            Dim o As Object = ViewState("SB_CnorTown")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorTown") = Value
        End Set
    End Property
    
    Property psCnorState() As String
        Get
            Dim o As Object = ViewState("SB_CnorState")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorState") = Value
        End Set
    End Property
    
    Property psCnorPostCode() As String
        Get
            Dim o As Object = ViewState("SB_CnorPostCode")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorPostCode") = Value
        End Set
    End Property
    
    Property psCnorCountryName() As String
        Get
            Dim o As Object = ViewState("SB_CnorCountryName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCountryName") = Value
        End Set
    End Property
    
    Property psCnorCountryKey() As String
        Get
            Dim o As Object = ViewState("SB_CnorCountryKey")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCountryKey") = Value
        End Set
    End Property
    
    Property psCnorCtcName() As String
        Get
            Dim o As Object = ViewState("SB_CnorCtcName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCtcName") = Value
        End Set
    End Property
    
    Property psCnorCtcTel() As String
        Get
            Dim o As Object = ViewState("SB_CnorCtcTel")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCtcTel") = Value
        End Set
    End Property
    
    Property psCnorCtcEmail() As String
        Get
            Dim o As Object = ViewState("SB_CnorCtcEmail")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CnorCtcEmail") = Value
        End Set
    End Property
    
    Property pnCategoryMode() As Integer
        Get
            Dim o As Object = ViewState("SB_CategoryMode")
            If o Is Nothing Then
                'Return ConfigLib.GetConfigItem_CategoryCount
                Return 2 ' should never reach here
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("SB_CategoryMode") = Value
        End Set
    End Property
    
    Property psServiceLevel() As String
        Get
            Dim o As Object = ViewState("SB_ServiceLevel")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_ServiceLevel") = Value
        End Set
    End Property
    
    Property pbInQuickMode() As Boolean
        Get
            Dim o As Object = ViewState("SB_InQuickMode")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_InQuickMode") = Value
        End Set
    End Property
    
    Property plCneeAddressKey() As Long
        Get
            Dim o As Object = ViewState("SB_CneeAddressKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("SB_CneeAddressKey") = Value
        End Set
    End Property
    
    Property plConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("SB_ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("SB_ConsignmentKey") = Value
        End Set
    End Property
    
    Property plPerCustomerConfiguration() As Long
        Get
            Dim o As Object = ViewState("SB_PerCustomerConfiguration")
            If o Is Nothing Then
                Return PER_CUSTOMER_CONFIGURATION_NONE
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("SB_PerCustomerConfiguration") = Value
        End Set
    End Property
    
    Property plCneeCountryKey() As Long
        Get
            Dim o As Object = ViewState("SB_CneeCountryKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("SB_CneeCountryKey") = Value
        End Set
    End Property
    
    Property psCategory() As String
        Get
            Dim o As Object = ViewState("SB_Category")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_Category") = Value
        End Set
    End Property
    
    Property psSubCategory() As String
        Get
            Dim o As Object = ViewState("SB_SubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_SubCategory") = Value
        End Set
    End Property
    
    Property psSubSubCategory() As String
        Get
            Dim o As Object = ViewState("SB_SubSubCategory")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_SubSubCategory") = Value
        End Set
    End Property
    
    Property psDisplayMode() As String
        Get
            Dim o As Object = ViewState("SB_DisplayMode")
            If o Is Nothing Then
                Return "_"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_DisplayMode") = Value
        End Set
    End Property
    
    Property pbAbleToViewGlobalAddressBook() As Boolean
        Get
            Dim o As Object = ViewState("SB_AbleToViewGlobalAddressBook")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_AbleToViewGlobalAddressBook") = Value
        End Set
    End Property
    
    Property pbAbleToEditGlobalAddressBook() As Boolean
        Get
            Dim o As Object = ViewState("SB_AbleToEditGlobalAddressBook")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_AbleToEditGlobalAddressBook") = Value
        End Set
    End Property
    
    Property pbUsingSharedAddressBook() As Boolean
        Get
            Dim o As Object = ViewState("SB_UsingSharedAddressBook")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_UsingSharedAddressBook") = Value
        End Set
    End Property
    
    Property pbSiteApplyMaxGrabs() As Boolean
        Get
            Dim o As Object = ViewState("SB_SiteApplyMaxGrabs")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_SiteApplyMaxGrabs") = Value
        End Set
    End Property
    
    Property pbSiteShowZeroStockBalances() As Boolean
        Get
            Dim o As Object = ViewState("SB_SiteShowZeroStockBalances")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_SiteShowZeroStockBalances") = Value
        End Set
    End Property

    Property pbCalendarManagement() As Boolean
        Get
            Dim o As Object = ViewState("SB_CalendarManagement")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_CalendarManagement") = Value
        End Set
    End Property

    Property pbCustomLetters() As Boolean
        Get
            Dim o As Object = ViewState("SB_CustomLetters")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_CustomLetters") = Value
        End Set
    End Property

    Property pbOnDemandProducts() As Boolean
        Get
            Dim o As Object = ViewState("SB_OnDemandProducts")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_OnDemandProducts") = Value
        End Set
    End Property
   
    Property pbZeroStockNotifications() As Boolean
        Get
            Dim o As Object = ViewState("SB_ZeroStockNotifications")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_ZeroStockNotifications") = Value
        End Set
    End Property
   
    Property pbShowNotes() As Boolean
        Get
            Dim o As Object = ViewState("SB_ShowNotes")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_ShowNotes") = Value
        End Set
    End Property
   
    Property pbMultipleAddressOrders() As Boolean
        Get
            Dim o As Object = ViewState("SB_MultipleAddressOrders")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_MultipleAddressOrders") = Value
        End Set
    End Property

    Property pbOrderAuthorisation() As Boolean
        Get
            Dim o As Object = ViewState("SB_OrderAuthorisation")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_OrderAuthorisation") = Value
        End Set
    End Property

    Property pbProductAuthorisation() As Boolean
        Get
            Dim o As Object = ViewState("SB_ProductAuthorisation")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_ProductAuthorisation") = Value
        End Set
    End Property
    
    Property psRetrievedAddress() As String
        Get
            Dim o As Object = ViewState("SB_RetrievedAddress")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_RetrievedAddress") = Value
        End Set
    End Property
    
    Property psOnDemandSessionGUID() As String
        Get
            Dim o As Object = ViewState("SB_PODSessionGUID")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_PODSessionGUID") = Value
        End Set
    End Property
    
    Property pbUsesCategories() As Boolean
        Get
            Dim o As Object = ViewState("SB_UsesCategories")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_UsesCategories") = Value
        End Set
    End Property
    
    Property psVirtualThumbFolder() As String
        Get
            Dim o As Object = ViewState("SB_VirtualThumbFolder")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_VirtualThumbFolder") = Value
        End Set
    End Property

    Property pnAddressPage() As Integer
        Get
            Dim o As Object = ViewState("SB_AddressPage")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("SB_AddressPage") = Value
        End Set
    End Property
    
    Property pnAddressVirtualItemCount() As Integer
        Get
            Dim o As Object = ViewState("SB_AddressVirtualItemCount")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("SB_AddressVirtualItemCount") = Value
        End Set
    End Property

    Property pbAuthorisationRequired() As Boolean
        Get
            Dim o As Object = ViewState("SB_AuthorisationRequired")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_AuthorisationRequired") = Value
        End Set
    End Property
    
    Property pnAuthoriser() As Integer
        Get
            Dim o As Object = ViewState("SB_Authoriser")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("SB_Authoriser") = Value
        End Set
    End Property
    
    Property pbBasketContainsWURSCriticalProducts() As Boolean
        Get
            Dim o As Object = ViewState("SB_BasketContainsWURSCriticalProducts")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("SB_BasketContainsWURSCriticalProducts") = Value
        End Set
    End Property
    
    Property psProductView() As String
        Get
            Dim sProductView As String
            Dim o As Object = ViewState("SB_ProductView")
            If o Is Nothing Then
                If Request.Cookies("SprintConfig") Is Nothing Then
                    If IsHysterOrYale() Or IsBNI() Then
                        sProductView = PRODUCT_VIEW_RICH
                    Else
                        sProductView = PRODUCT_VIEW_CLASSIC
                    End If
                    Call CreateSprintConfigCookie(sProductView)
                Else
                    sProductView = Request.Cookies("SprintConfig")("SB_ProductView") & ""
                    If sProductView = String.Empty Then
                        If IsHysterOrYale() Or IsBNI() Then
                            sProductView = PRODUCT_VIEW_RICH
                        Else
                            sProductView = PRODUCT_VIEW_CLASSIC
                        End If
                        Call UpdateSprintConfigCookieProductView(sProductView)
                    End If
                End If
                ViewState("SB_ProductView") = sProductView
            Else
                sProductView = CStr(o)
            End If
            Return sProductView
        End Get
    
        Set(ByVal Value As String)
            Call UpdateSprintConfigCookieProductView(Value)
            ViewState("SB_ProductView") = Value
            UpdateSprintConfigCookieProductView(Value)
            lnkbtnDisplayModeChange.Text = Value
        End Set
    End Property
    
    Property pdblBasketTotalValue() As Double
        Get
            Dim o As Object = ViewState("SB_BasketTotalValue")
            If o Is Nothing Then
                Return 0.0
            End If
            Return CDbl(o)
        End Get
        Set(ByVal Value As Double)
            ViewState("SB_BasketTotalValue") = Value
        End Set
    End Property
    
    Property plBasketWeightGrams() As Long
        Get
            Dim o As Object = ViewState("SB_BasketWeightGrams")
            If o Is Nothing Then
                Return 0
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("SB_BasketWeightGrams") = Value
        End Set
    End Property
    
    Property pdtCalendarManagedSelectionX() As Date
        Get
            Dim o As Object = ViewState("SB_CalendarManagedSelectionX")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CDate(o)
        End Get
        Set(ByVal Value As Date)
            If Value = Nothing Then
                lblCMDeliveryDate.Text = String.Empty
            Else
                lblCMDeliveryDate.Text = Value.ToString("dd-MMM-yy")
            End If
            ViewState("SB_CalendarManagedSelectionX") = Value
        End Set
    End Property
    
    Property pdtCalendarManagedSelectionY() As Date
        Get
            Dim o As Object = ViewState("SB_CalendarManagedSelectionY")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CDate(o)
        End Get
        Set(ByVal Value As Date)
            If Value = Nothing Then
                lblCMCollectionDate.Text = String.Empty
            Else
                lblCMCollectionDate.Text = Value.ToString("dd-MMM-yy")
            End If
            ViewState("SB_CalendarManagedSelectionY") = Value
        End Set
    End Property
    
    Property pdictCurrentMonthBookings() As Dictionary(Of Date, String)
        Get
            Dim o As Object = ViewState("SB_Bookings")
            If o Is Nothing Then
                Return Nothing
            End If
            Return CType(o, Dictionary(Of Date, String))
        End Get
        Set(ByVal Value As Dictionary(Of Date, String))
            ViewState("SB_Bookings") = Value
        End Set
    End Property

    Property psCustomLetterText() As String
        Get
            Dim o As Object = ViewState("SB_CustomLetterText")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CustomLetterText") = Value
        End Set
    End Property

    Property psCustomLetterInstructions() As String
        Get
            Dim o As Object = ViewState("SB_CustomLetterInstructions")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("SB_CustomLetterInstructions") = Value
        End Set
    End Property

    Protected Function InvoiceAddressRecordExists() As Boolean
        Dim sSQL As String = "SELECT id FROM OnDemandTransaction WHERE SessionGUID = '" & psOnDemandSessionGUID & "'"
        Dim oDataTable As DataTable
        oDataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count > 0 Then
            InvoiceAddressRecordExists = True
        Else
            InvoiceAddressRecordExists = False
        End If
    End Function
    
    Protected Sub btnBackInvoiceAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowDeliveryAddressPanel()
    End Sub
    
    Protected Function RetrievePage(ByVal sURL As String) As String
        RetrievePage = String.Empty
        Dim wr As System.Net.WebRequest
        Try
            wr = WebRequest.Create(sURL)
            Dim resp As HttpWebResponse = wr.GetResponse()
            Dim sr As New StreamReader(resp.GetResponseStream)
            RetrievePage = sr.ReadToEnd()
            sr.Close()
        Catch ex As Exception
            WebMsgBox.Show("Error in RetrievePage: (" & sURL & " ): " & ex.Message)
        End Try
    End Function
    
    Protected Sub ddlItemsPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        gvProductList.PageSize = ddl.SelectedValue
        gvProductList.PageIndex = 0
        Call BindProductGridDispatcher("ProductCode")
    End Sub
    
    Protected Sub lnkbtnCMAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkbtnCMFindEventAddress.Visible = False
        Call ShowCMCountryDropdowns()
    End Sub

    Protected Sub lnkbtnCMCollectionAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkbtnCMFindCollectionAddress.Visible = False
        Call ShowCMCountryDropdowns()
    End Sub
    
    Protected Sub ShowCMCountryDropdowns()
        trCMCountry.Visible = True
        lnkbtnCMAddressOutsideUK.Visible = False
        'trCMCollectionCountry.Visible = True
        lnkbtnCMCollectionAddressOutsideUK.Visible = False
        Call InitCMCountryDropdowns()
    End Sub
    
    Protected Sub InitCMCountryDropdowns()
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
    End Sub

    Protected Sub lnkbtnCMAddSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'trCMContactName2.Visible = True
        lblLegendCMContactName2.Visible = True
        rfvCMContactName2.Visible = True
        tbCMContactName2.Visible = True
        trCMContactPhone2.Visible = True
        trCMContactMobile2.Visible = True
        lnkbtnCMAddSecondContact.Visible = False
        tbCMContactName2.Focus()
    End Sub

    Protected Sub lnkbtnFindAvailableProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowFindAvailableProductsPanel()
    End Sub
    
    Protected Sub btnBackFromFindAvailableProductsToEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        Call ShowCalendarManagedPanel(bSkipClearDateSelection:=True)
    End Sub
    
    Protected Sub btnCMAddAvailableProductsToBasket_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If cbCMRemoveNonBookableProductsFromBasket.Checked Then
            Call GetBasketFromSession()
            For i As Integer = (gdtBasket.Rows.Count - 1) To 0 Step -1
                Dim dr As DataRow = gdtBasket.Rows(i)
                If CBool(dr("CalendarManaged")) Then
                    If dr("Notes").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN1) Or dr("ProductDate").ToString.ToLower.Contains(CALENDAR_MANAGED_NON_BOOKABLE_TOKEN2) Then
                        gdtBasket.Rows(i).Delete()
                        Session(gsBasketCountName) = Session(gsBasketCountName) - 1
                    End If
                End If
            Next
            Call SaveBasketToSession()
        End If
        For Each gvr As GridViewRow In gvCMAvailableProducts.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim cb As CheckBox = gvr.Cells(0).FindControl("cbCMAddAvailableProductToBasket")
                If cb.Checked Then
                    Dim hid As HiddenField = gvr.Cells(0).FindControl("hidAvailableProductKey")
                    Call AddItemToBasket(hid.Value, bIsFromCookieBasket:=False)
                End If
            End If
        Next
        Call HideAllPanels()
        Call ShowCalendarManagedPanel(bSkipClearDateSelection:=True)
    End Sub
    
    Protected Sub lnkbtnCneeCountryUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Integer = 1 To ddlCneeCountry.Items.Count - 1
            If ddlCneeCountry.Items(i).Text = "U.K." Or ddlCneeCountry.Items(i).Text = "UK" Then
                ddlCneeCountry.SelectedIndex = i
                Call SetCountry(ddlCneeCountry.SelectedValue, "")
                txtCneeName.Focus()
                Exit For
            End If
        Next
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
        txtCneeState.Visible = True
        lblLegendRegion.Text = "County / Region"
        lblLegendRegion.ForeColor = Drawing.Color.Blue
        txtCneeState.Text = String.Empty
        lblLegendRegion.Font.Bold = False
        rfvRegion.Enabled = False
        lblLegendPostcodeZipcode.Text = "Post Code:"
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
        txtCneeState.Text = String.Empty
        lblLegendRegion.Font.Bold = True
        lblLegendPostcodeZipcode.Text = "Zip Code:"
    End Sub
    
    Protected Sub SetCountryUSANewYorkCity()
        Call HideCountryRelatedControls()
        lblLegendNewYorkCity.Visible = True
        lblLegendRegion.Text = "State"
        lblLegendRegion.ForeColor = Drawing.Color.Red
        rfvRegion.Enabled = False
        txtCneeState.Text = lblLegendNewYorkCity.Text
        lblLegendPostcodeZipcode.Text = "Zip Code:"
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
        txtCneeState.Text = String.Empty
        lblLegendRegion.Font.Bold = True
        lblLegendPostcodeZipcode.Text = "Postal Code:"
    End Sub
    
    Protected Sub HideCountryRelatedControls()
        ddlUSStatesCanadianProvinces.Visible = False
        lblLegendNewYorkCity.Visible = False
        txtCneeState.Visible = False
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

    Protected Sub ddlCneeCountry_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        Call SetCountry(ddlCneeCountry.SelectedValue, "")
        txtCneeName.Focus()
    End Sub

    Protected Sub ddlUSStatesCanadianProvinces_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        If ddlUSStatesCanadianProvinces.SelectedIndex > 0 Then
            txtCneeState.Text = ddlUSStatesCanadianProvinces.SelectedItem.Text
        Else
            txtCneeState.Text = String.Empty
        End If
    End Sub
    
    Protected Sub lnkbtnRamblersGroup_Click(sender As Object, e As System.EventArgs)
        Dim lb As LinkButton = sender
        xdsVar26RamblersAreaGroups.XPath = "RamblersAreaGroups/Ramblers" & lb.CommandArgument & "/areaGroup"
        'ddlRamblersAreaGroup.Items.Insert(0, New ListItem("- select area/group -"))
        xdsVar26RamblersAreaGroups.DataBind()
        'ddlRamblersAreaGroup.DataBind()
    End Sub
    
    Protected Sub ddlPerCustomerConfiguration30PrintServiceLevel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim dblTotalPrice As Double = 0.0
        If ddl.SelectedIndex > 0 Then
            Call GetBasketFromSession()
            For Each dr As DataRow In gdtBasket.Rows
                'Dim nLogisticProductKey As Int32 = dr("ProductKey")
                Dim nQtyToPick As Int32 = dr("QtyToPick")
                Dim nPrintType As Int32 = ExecuteQueryToDataTable("SELECT Misc2 FROM LogisticProduct WHERE LogisticProductKey = " & dr("ProductKey")).Rows(0).Item(0)
                Dim drPrintCostMatrix As DataRow = ExecuteQueryToDataTable("SELECT * FROM ClientData_Jupiter_PrintCost WHERE [id] = " & nPrintType).Rows(0)
                Dim sPriceColumn As String
                If nQtyToPick <= 50 Then
                    sPriceColumn = "50"
                ElseIf nQtyToPick <= 100 Then
                    sPriceColumn = "100"
                ElseIf nQtyToPick <= 250 Then
                    sPriceColumn = "250"
                ElseIf nQtyToPick <= 500 Then
                    sPriceColumn = "500"
                ElseIf nQtyToPick <= 1000 Then
                    sPriceColumn = "1000"
                Else
                    WebMsgBox.Show("Print quantity error - please inform development")
                    Exit Sub
                End If
                sPriceColumn = "Price" & ddl.SelectedValue & "UpTo" & sPriceColumn
                Dim dblPrice As Double = drPrintCostMatrix(sPriceColumn)
                dblTotalPrice += dblPrice
            Next
        Else
        End If
        lblPerCustomerConfiguration30TotalPrintCost.Text = "£" & Format(dblTotalPrice, "##,##0.00")
        lblPerCustomerConfiguration30PrintServiceLevel.Text = ddl.SelectedItem.Text
        lblPerCustomerConfiguration30ConfirmationTotalPrintCost.Text = "£" & Format(dblTotalPrice, "##,##0.00")
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Place an Order</title>
    <style type="text/css">
        .NoUnderline:
        {
            text-decoration: none;
        }
    </style>
</head>
<body>
    <form id="frmOrder" runat="Server">
    <main:Header ID="ctlHeader" runat="server"></main:Header>
    <table id="tblHeading" style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_order">
            <td style="width: 50%; white-space: nowrap; height: 15px;">
                &nbsp;&nbsp;<asp:Label ID="Label17" runat="server" ForeColor="White" Font-Size="XX-Small"
                    Font-Names="Verdana">Click</asp:Label>
                &nbsp;<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="ConditionsOfCarriage.pdf"
                    TabIndex="-1" ForeColor="#F9D938" Font-Size="XX-Small" Font-Names="Verdana" Target="_blank">here</asp:HyperLink>
                &nbsp;<asp:Label ID="Label8" runat="server" ForeColor="White" Font-Size="XX-Small"
                    Font-Names="Verdana">to see our conditions of carriage</asp:Label>
            </td>
            <td style="width: 50%; white-space: nowrap; height: 15px;" align="right">
                <asp:Label ID="lblBasketMsg" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    ForeColor="White">Your basket contains</asp:Label>
                &nbsp;<asp:Label runat="server" ID="lblBasketCount" ForeColor="#F9D938" Font-Bold="true"
                    Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                &nbsp;<asp:Label ID="lblBasketItemPlural" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                    ForeColor="White" Text="items" />&nbsp;&nbsp;
            </td>
        </tr>
    </table>
    <table id="tblTopButtons" style="width: 100%; height: 26px; font-family: Verdana;
        font-size: x-small" cellpadding="0" cellspacing="0">
        <tr valign="middle">
            <td align="left" valign="middle" style="white-space: nowrap">
                <asp:Button ID="btnShowByCategory" runat="server" OnClick="btnShowByCategory_click"
                    Text="show categories" ToolTip="show products by category" CausesValidation="false" />
                &nbsp;&nbsp;<asp:Button ID="btn_ShowFullProdList" runat="server" OnClick="btn_ShowFullProdList_click"
                    Text="show all products" ToolTip="get full product list" CausesValidation="false" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label16" runat="server" ForeColor="Gray"
                    Font-Size="XX-Small" Font-Names="Verdana">search:</asp:Label>
                &nbsp;<asp:TextBox runat="server" Width="80px" Font-Size="XX-Small" Font-Names="Verdana"
                    ID="txtProdSearchCriteria" MaxLength="50"></asp:TextBox>
                &nbsp;<asp:Button ID="btn_SearchProd" runat="server" OnClick="btn_SearchProd_click"
                    Text="go" ToolTip="search across all products" CausesValidation="false" />
            </td>
            <td align="right" valign="middle" style="white-space: nowrap">
                <asp:LinkButton ID="lnkbtnDisplayModeChange" runat="server" ForeColor="Gray" Font-Names="Verdana"
                    Font-Size="XX-Small" Text="not initialised" CausesValidation="false" OnClick="lnkbtnDisplayModeChange_Click" />&nbsp;&nbsp;
                <span id="spanQuickModeCheckBox" visible="true" runat="server">
                    <asp:CheckBox ID="chk_QuickMode" runat="server" Text="quick mode" ForeColor="Gray"
                        OnCheckedChanged="chk_QuickMode_CheckedChanged" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="true"></asp:CheckBox>&nbsp;&nbsp;<a onmouseover="return escape('<b>quick mode</b> allows you to select several products without returning to your basket between each selection')"
                            style="color: silver; cursor: help">&nbsp;?&nbsp;</a>&nbsp;</span> &nbsp;<asp:Button
                                ID="btn_viewbasket" runat="server" OnClick="btn_ViewCurrentBasket_click" Text="view basket"
                                ToolTip="view current order" CausesValidation="false" />
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlCategorySelection1" runat="server" Visible="True" Width="100%">
        <table id="tblCategorySelection" runat="server" width="100%" style="font-family: Verdana;
            font-size: small" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 2%">
                </td>
                <td valign="top" style="white-space: nowrap; width: 48%; background-color: #DFE5E2">
                    &nbsp;&nbsp;<asp:Label ID="Label111" runat="server" ForeColor="Navy" Font-Bold="True"
                        Font-Size="X-Small">Product Categories</asp:Label>
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
                        Font-Bold="True" Font-Size="X-Small">Sub Categories</asp:Label>
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
                        Font-Bold="True" Font-Size="X-Small">Sub Category 1</asp:Label>
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
                        Font-Bold="True" Font-Size="X-Small">Sub Category 2</asp:Label>
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
        &nbsp;&nbsp;<asp:Label BackColor="#F9D938" runat="server" ID="lblCategoryHeader"
            Font-Names="Verdana" ForeColor="Navy" Font-Size="X-Small"></asp:Label>
        &nbsp;&nbsp;<asp:Label ID="lblProductMessage" BackColor="#F9D938" runat="server"
            ForeColor="Navy" Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
        <asp:GridView ID="gvProductList" runat="server" PageSize="5" Width="100%" Font-Names="Verdana"
            Font-Size="XX-Small" Visible="False" AllowPaging="True" AutoGenerateColumns="False"
            GridLines="None" ShowFooter="True" OnPageIndexChanging="gvProductList_PageIndexChanging"
            OnRowDataBound="gvProductList_RowDataBound">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <PagerSettings Mode="NumericFirstLast" FirstPageText="First" LastPageText="Last"
                NextPageText="Next" PreviousPageText="Prev" Position="Top" />
            <Columns>
                <asp:BoundField DataField="LogisticProductKey" Visible="False" />
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:HiddenField ID="hidProductkey" runat="server" Value='<%# Eval("LogisticProductKey")%>' />
                        <asp:HiddenField ID="hidCustomLetterRichView" runat="server" Value='<%# Eval("CustomLetter")%>' />
                        <asp:HiddenField ID="hidNotes" runat="server" Value='<%# Eval("Notes")%>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField>
                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <table id="tabAuthorisationInfo" runat="server" visible='<%# bSetAuthorisationInfoVisibility(Container.DataItem) %>'
                            width="100%" style="font-family: Verdana; font-size: xx-small; color: Gray" border="0">
                            <tr>
                                <td style="width: 2%">
                                    &nbsp;
                                </td>
                                <td style="width: 40%">
                                    &nbsp;
                                </td>
                                <td align="right" style="width: 56%">
                                    <asp:Label ID="lblAuthorisationMessage" runat="server" ForeColor="red" Text='<%# sSetAuthorisationInfo(Container.DataItem) %>' />
                                </td>
                                <td style="width: 2%">
                                    &nbsp;
                                </td>
                            </tr>
                        </table>
                        <table width="100%" style="font-family: Verdana; font-size: xx-small; color: Gray"
                            border="0">
                            <tr>
                                <td id="tdMiscRowSpan1" runat="server" rowspan="5" style="width: 2%">
                                    &nbsp;
                                </td>
                                <td id="tdMiscRowSpan2" runat="server" rowspan="4" valign="top" style="width: 10%">
                                    <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'></asp:HyperLink>
                                </td>
                                <td valign="top" style="width: 11%">
                                    <asp:Label ID="Label10" runat="server" ForeColor="Gray">Product&nbsp;Code:</asp:Label>
                                </td>
                                <td valign="top" style="width: 20%">
                                    <asp:Label ID="Label11" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductCode") %></asp:Label>
                                </td>
                                <td valign="top" style="width: 11%">
                                    <asp:Label ID="Label21" runat="server" ForeColor="Gray"><%# gvProductListGetLegend("Product Date:") %></asp:Label>
                                </td>
                                <td valign="top" style="width: 23%">
                                    <asp:Label ID="Label22" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductDate") %></asp:Label>
                                </td>
                                <td valign="top" align="right" style="width: 21%">
                                    <asp:Label ID="Label2300" runat="server" ForeColor="Gray" Text="You can order" />
                                    &nbsp;<asp:Label ID="Label24" runat="server" ForeColor="Navy"><%# Format(DataBinder.Eval(Container.DataItem,"Quantity"),"#,##0") %></asp:Label>
                                </td>
                                <td style="width: 2%">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr id="trCategorySubcategory" runat="server" visible="<%# pbUsesCategories %>">
                                <td valign="top">
                                    <asp:Label ID="Label25" runat="server" ForeColor="Gray">Category:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label26" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"ProductCategory") %></asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label27" runat="server" ForeColor="Gray">Sub&nbsp;Category:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label28" runat="server" ForeColor="Navy" Text='<%# gvProductListSetSubCategory(Container.DataItem) %>' />
                                </td>
                                <td align="right">
                                    <asp:Label ID="lblCostCentreLegend" runat="server" ForeColor="Gray" Text="Cost Ctr/Dept: " /><asp:Label
                                        ID="lblCostCentre" runat="server" ForeColor="Navy" Text='<%# DataBinder.Eval(Container.DataItem,"ProductDepartmentId") %>' />
                                </td>
                            </tr>
                            <%--                                <tr id="trMiscLine1" runat="server" visible="false">
                                    <td valign="top">
                                        <asp:Label ID="Label25a" runat="server" forecolor="Gray">Promo&nbsp;Stuff&nbsp;Type:</asp:Label>
                                    </td>
                                    <td valign="top">
                                        <asp:Label ID="Label26a" runat="server" forecolor="Navy"><%# DataBinder.Eval(Container.DataItem,"Misc1") %></asp:Label>
                                    </td>
                                    <td valign="top">
                                        <asp:Label ID="Label27a" runat="server" forecolor="Gray">NAH&nbsp;Medicine:</asp:Label>
                                    </td>
                                    <td valign="top">
                                        <asp:Label ID="Label28a" runat="server" forecolor="Navy" Text='<%# DataBinder.Eval(Container.DataItem,"Misc2") %>'/>
                                    </td>
                                    <td align="right">
                                        &nbsp;
                                    </td>
                                </tr> --%>
                            <tr id="trQuantumLeap" runat="server" visible="<%# IsQuantumLeap() %>">
                                <td valign="top">
                                    <asp:Label ID="Label25a" runat="server" ForeColor="Gray">Supplier:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label26a" runat="server" ForeColor="Navy" Text='<%# DataBinder.Eval(Container.DataItem,"Misc1") %>' />
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label27a" runat="server" ForeColor="Gray">Boxed&nbsp;to&nbsp;Ship:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label28a" runat="server" ForeColor="Navy" Text='<%# DataBinder.Eval(Container.DataItem,"Misc2") %>' />
                                </td>
                                <td align="right">
                                    <asp:Label ID="Labeqll230" runat="server" ForeColor="Gray">SP:</asp:Label>&nbsp;
                                    <asp:Label ID="Labelql101" runat="server" ForeColor="Navy"><%#DataBinder.Eval(Container.DataItem, "UnitValue2", "{0:c}")%></asp:Label>
                                </td>
                            </tr>
                            <tr style="height: 15px" id="trProdOwnerLine1" runat="server" visible="false">
                                <td valign="top">
                                    <asp:Label ID="Label25ab" runat="server" ForeColor="Gray">Stock Owner 1:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label26ab" runat="server" ForeColor="Navy" Text='<%# sSetProductOwnerInfo(DataBinder.Eval(Container.DataItem,"ProductOwner1")) %>' />
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label27ab" runat="server" ForeColor="Gray">Stock Owner 2:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label28ab" runat="server" ForeColor="Navy" Text='<%# sSetProductOwnerInfo(DataBinder.Eval(Container.DataItem,"ProductOwner2")) %>' />
                                </td>
                                <td align="right">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <asp:Label ID="Label40" runat="server" ForeColor="Gray" Visible="<%# IsNotQuantumLeap() %>"
                                        Text="Language:" />
                                    <asp:Label ID="Label40bis" runat="server" ForeColor="Gray" Visible="<%# IsQuantumLeap() %>"
                                        Text="Barcode:" />
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label41" runat="server" ForeColor="Navy"><%#DataBinder.Eval(Container.DataItem, "LanguageID")%></asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="lblLegendQtyPerBox" runat="server" Visible="<%# IsNotStruttWUHysterYaleCAB() %>"
                                        ForeColor="Gray">Qty/box:</asp:Label>
                                </td>
                                <td valign="top">
                                    <asp:Label ID="Label56" runat="server" Visible="<%# IsNotStruttWUHysterYaleCAB() %>"
                                        ForeColor="Navy"><%#DataBinder.Eval(Container.DataItem, "ItemsPerBox")%></asp:Label>
                                </td>
                                <td align="right">
                                    <asp:Label ID="Label230" runat="server" Visible="<%# IsNotWURS() %>" ForeColor="Gray">Value:</asp:Label>&nbsp;
                                    <asp:Label ID="Label410" runat="server" ForeColor="Navy" Visible="<%# IsNotWURS() %>"><%#DataBinder.Eval(Container.DataItem, "UnitValue", "{0:c}")%></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <asp:Label ID="Label29" runat="server" ForeColor="Gray">Description:</asp:Label>
                                </td>
                                <td valign="top" colspan="3" rowspan="2">
                                    <asp:Label ID="Label30" runat="server" ForeColor="Navy" Font-Bold="<%# IsJupiter() %>"><%# DataBinder.Eval(Container.DataItem,"ProductDescription") %></asp:Label>
                                </td>
                                <td valign="bottom" align="right" rowspan="2">
                                    <asp:Button ID="btn_CMShowUsage" runat="server" Text="calendar" Visible='<%# gvProductListShowUsage(Container.DataItem) %>'
                                        ToolTip="show usage" OnClientClick='<%# "Javascript:CMShowUsage(""" & DataBinder.Eval(Container.DataItem,"LogisticProductKey") & """)" %> ' />
                                    <asp:Button ID="btn_AddToBasket" runat="server" Text="add to basket" Visible='<%# gvProductListSetAddToBasketVisibility(Container.DataItem) %>'
                                        ToolTip="add product to current order" OnClick="btn_AddToBasket_Click" />
                                    <asp:Button ID="btnZeroStockNotification" runat="server" Text="notify me" ToolTip="notify me when this product is in stock"
                                        Visible='<%# gvProductListSetZeroStockNotificationVisibility(Container.DataItem) %>'
                                        OnClick="btnZeroStockNotification_Click" />
                                    <asp:Label ID="lblProductStatusMessage" runat="server" ForeColor="Red"><%# gvProductListGetProductStatusMessage() %></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <asp:HyperLink runat="server" ID="hlnk_PDF" Target="_blank" ForeColor="blue" ToolTip="click here to view a PDF document for this product">view&nbsp;pdf</asp:HyperLink>
                                </td>
                            </tr>
                        </table>
                        <table id="tblProductNotes" runat="server" visible="true" width="100%" style="font-family: Verdana;
                            font-size: xx-small; color: Gray">
                            <tr>
                                <td style="width: 12%">
                                    &nbsp;
                                </td>
                                <td style="width: 20%">
                                    <asp:Label ID="Label129" runat="server" ForeColor="Gray">Notes:</asp:Label>
                                </td>
                                <td style="width: 76%">
                                    <asp:Label ID="lblProductNotes" ForeColor="red" runat="server" Font-Bold="<%# IsJupiter() %>"><%#DataBinder.Eval(Container.DataItem, "Notes")%></asp:Label>
                                </td>
                                <td style="width: 2%">
                                </td>
                            </tr>
                        </table>
                        <table width="100%" style="font-family: Verdana; font-size: xx-small; color: Gray">
                            <tr>
                                <td colspan="7" valign="top">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#E0E0E0" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy"
                HorizontalAlign="Center" />
        </asp:GridView>
        <asp:Label ID="Label126" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Items per page:"></asp:Label>&nbsp;<asp:DropDownList ID="ddlItemsPerPage" runat="server"
                AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlItemsPerPage_SelectedIndexChanged">
                <asp:ListItem Selected="True">5</asp:ListItem>
                <asp:ListItem>20</asp:ListItem>
                <asp:ListItem>50</asp:ListItem>
            </asp:DropDownList>
        <asp:Table ID="Table3" runat="Server" Width="100%" Font-Size="X-Small" Font-Names="Verdana">
            <asp:TableRow>
                <asp:TableCell></asp:TableCell>
                <asp:TableCell HorizontalAlign="Right">
                    <asp:LinkButton runat="server" ForeColor="Blue" ID="btnRefreshProductList" Font-Size="XX-Small"
                        Font-Names="Verdana" Visible="False" OnClick="btn_RefreshProdList_click" ToolTip="refresh current stock levels for fast moving products">refresh product list</asp:LinkButton>&nbsp;&nbsp;
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlBasket" runat="server" Visible="False" Width="100%">
        <asp:Table ID="tabBasketHeader" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Width="100%" CellSpacing="1" CellPadding="2">
            <asp:TableRow>
                <asp:TableCell Width="5%"></asp:TableCell>
                <asp:TableCell Width="45%"></asp:TableCell>
                <asp:TableCell Width="45%"></asp:TableCell>
                <asp:TableCell Width="5%"></asp:TableCell>
            </asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell>
                <asp:TableCell ColumnSpan="2">
                    <asp:Label ID="Label18" runat="server" Font-Size="Small" ForeColor="Navy">Basket</asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="Label15" runat="server" Font-Size="XX-Small" ForeColor="Navy" Text="Specify a quantity for each product before proceeding to checkout.
                        At the next stage you will enter the delivery address and provide any special instructions required for this order." />
                    <span id="spanMultiAddressInstructions" runat="server" visible="false">
                        <br />
                        <br />
                        For orders to multiple destinations proceed to checkout, select the destination
                        list from the dropdown box, then specify the quantity required for each destination.</span>
                </asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell>
                    <br />
                    <asp:Button ID="Button1" runat="server" OnClick="btn_ReturnToProducts_click" Text="return to products"
                        ToolTip="go back to browse the stock items" />
                </asp:TableCell><asp:TableCell HorizontalAlign="Right">
                    <br />
                    <asp:Button ID="Button2" runat="server" OnClick="btn_CheckOut_click" Text="checkout"
                        ToolTip="proceed to checkout" />
                </asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell ColumnSpan="2">
                    <asp:Label ID="lblBasketMessage1" runat="server" BackColor="#F9D938" ForeColor="Navy"
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                </asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
        </asp:Table>
        <asp:DataGrid ID="gvBasket" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"
            Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True"
            OnItemCommand="gvBasket_item_click">
            <FooterStyle Wrap="False"></FooterStyle>
            <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
            <Columns>
                <asp:BoundColumn Visible="False" DataField="ProductKey">
                    <ItemStyle Wrap="False"></ItemStyle>
                </asp:BoundColumn>
                <asp:TemplateColumn>
                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    <ItemTemplate>
                        <table width="100%" style="font-family: Verdana; font-size: xx-small; color: Gray"
                            border="0">
                            <tr>
                                <td style="width: 5%">
                                    &nbsp;
                                </td>
                                <td style="width: 10%">
                                    &nbsp;
                                </td>
                                <td style="width: 12%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label7" runat="server" ForeColor="Gray">Product Code:</asp:Label>
                                </td>
                                <td style="width: 20%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label8" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductCode") %></asp:Label>
                                </td>
                                <td style="width: 12%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label9" runat="server" ForeColor="Gray">Version/Date:</asp:Label>
                                </td>
                                <td style="width: 20%; white-space: nowrap" valign="top">
                                    <asp:Label ID="Label12" runat="server" ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductDate") %></asp:Label>
                                </td>
                                <td style="width: 16%; white-space: nowrap" align="right" valign="top">
                                    <asp:Label ID="Label13" runat="server" ForeColor="Gray">Qty Available:</asp:Label>&nbsp;
                                    <asp:Label ID="lblQtyAvailable" runat="server" ForeColor="Navy"><%#Format(DataBinder.Eval(Container.DataItem, "QtyAvailable"), "#,##0")%></asp:Label><asp:HiddenField
                                        ID="hidQtyAvailable" Value='<%# DataBinder.Eval(Container.DataItem,"QtyAvailable") %>'
                                        runat="server" />
                                    <asp:HiddenField ID="hidCalendarManaged" Value='<%# DataBinder.Eval(Container.DataItem,"CalendarManaged") %>'
                                        runat="server" />
                                    <asp:HiddenField ID="hidCustomLetter" Value='<%# DataBinder.Eval(Container.DataItem,"CustomLetter") %>'
                                        runat="server" />
                                </td>
                                <td style="width: 5%">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td rowspan="4">
                                    &nbsp;
                                </td>
                                <td rowspan="4">
                                    <asp:HyperLink ID="HyperLink3" runat="server" ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'
                                        NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ToolTip="click here to see larger image" />
                                </td>
                                <td style="white-space: nowrap">
                                    <asp:Label ID="Label17" runat="server" ForeColor="Gray">Language:</asp:Label><asp:Label
                                        ID="Label20" runat="server" ForeColor="Red" Text='<%# DataBinder.Eval(Container.DataItem,"LanguageID") %>' />
                                </td>
                                <td style="white-space: nowrap">
                                    <asp:Label ID="Label18" runat="server" Visible="<%# IsNotStruttWUHysterYaleCAB() %>"
                                        ForeColor="Gray">Qty/box:</asp:Label>
                                </td>
                                <td style="white-space: nowrap">
                                    <asp:Label ID="Label19" runat="server" Visible="<%# IsNotStruttWUHysterYaleCAB() %>"
                                        ForeColor="Red"><%# DataBinder.Eval(Container.DataItem,"BoxQty") %></asp:Label>
                                </td>
                                <td style="white-space: nowrap" align="right" valign="top">
                                </td>
                                <td rowspan="4">
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="Label16" runat="server" ForeColor="Gray">Description:</asp:Label>
                                </td>
                                <td colspan="3" rowspan="3" valign="top">
                                    <asp:Label ID="Label25" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"Description") %></asp:Label>
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
                                    &nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td style="white-space: nowrap">
                                    <asp:LinkButton ID="LinkButton2" runat="server" CommandName="Remove" ForeColor="Blue"
                                        ToolTip="remove this item from your order">remove item</asp:LinkButton>
                                </td>
                                <td style="white-space: nowrap" align="right">
                                    <asp:Label ID="lblLegendOrderQuantity" runat="server" Font-Size="XX-Small">Order Quantity:</asp:Label><asp:TextBox
                                        ID="txtOrderQuantity" runat="server" Font-Size="XX-Small" ForeColor="Navy" MaxLength="6"
                                        Text='<%# DataBinder.Eval(Container.DataItem,"QtyToPick") %>' Width="50px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    &nbsp;
                                </td>
                                <td colspan="6" valign="top">
                                    <hr />
                                </td>
                                <td>
                                    &nbsp;
                                </td>
                            </tr>
                        </table>
                    </ItemTemplate>
                </asp:TemplateColumn>
            </Columns>
        </asp:DataGrid><table id="tabBasketFooter" style="font-family: Verdana; font-size: xx-small;
            width: 100%">
            <tr>
                <td style="width: 5%; height: 26px;">
                </td>
                <td style="width: 45%; height: 26px;">
                    <asp:Button ID="ReturnToProducts_click" runat="server" OnClick="btn_ReturnToProducts_click"
                        Text="return to products" ToolTip="go back to browse the stock items" />
                </td>
                <td style="width: 45%; height: 26px;" align="right">
                    <asp:Button ID="btn_CheckOut" runat="server" OnClick="btn_CheckOut_click" Text="checkout"
                        ToolTip="proceed to checkout" />
                </td>
                <td style="width: 5%; height: 26px;">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblBasketMessage2" runat="server" BackColor="#F9D938" ForeColor="Navy"
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlAssociatedProducts" runat="server" Visible="False" Width="100%">
            <asp:Table ID="tabAssocProdsHeader" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                Width="100%" CellSpacing="1" CellPadding="2">
                <asp:TableRow runat="server">
                    <asp:TableCell Width="5%" runat="server"></asp:TableCell><asp:TableCell Width="90%"
                        runat="server">
                            <br/>
                            <asp:Label runat="server" BackColor="#F9D938" Font-Size="X-Small" ForeColor="Navy"> Associated Products - please review before proceeding </asp:Label>

                            <br/>
                            <br/>
                            <asp:Label runat="server" Font-Size="XX-Small">One or more items in your basket has had another product(s) associated with
                            it - see below. Please view these items to ensure your order is complete. Click the 'add to basket' button to add the associated
                            product to your basket above.</asp:Label>
                    </asp:TableCell><asp:TableCell Width="5%" runat="server"></asp:TableCell></asp:TableRow>
            </asp:Table>
            <asp:GridView ID="gvAssocProducts" runat="server" Width="100%" Font-Names="Verdana"
                Font-Size="XX-Small" Visible="False" AutoGenerateColumns="False" GridLines="None"
                ShowFooter="True" OnRowDataBound="gvAssocProducts_RowDataBound">
                <FooterStyle Wrap="False"></FooterStyle>
                <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
                <Columns>
                    <asp:BoundField DataField="ProductKey" Visible="false" />
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:HiddenField ID="hidAssProdkey" runat="server" Value='<%# Eval("ProductKey")%>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        <ItemTemplate>
                            <asp:Table ID="tabAssocProds" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                Width="100%" ForeColor="Gray">
                                <asp:TableRow>
                                    <asp:TableCell RowSpan="3" Width="5%"></asp:TableCell><asp:TableCell RowSpan="3"
                                        Width="10%" VerticalAlign="Top">
                                        <asp:HyperLink ID="hlnk_ThumbNail" runat="server" ToolTip="click here to see larger image"
                                            NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                            ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'></asp:HyperLink>
                                    </asp:TableCell><asp:TableCell Width="12%" VerticalAlign="Top" Wrap="False">
                                            <asp:Label runat="server" forecolor="Gray">Product Code:</asp:Label>
                                    </asp:TableCell><asp:TableCell Width="20%" VerticalAlign="Top" Wrap="False">
                                            <asp:Label runat="server" forecolor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductCode") %></asp:Label>
                                    </asp:TableCell><asp:TableCell Width="12%" VerticalAlign="Top" Wrap="False">
                                            <asp:Label runat="server" forecolor="Gray">Product Date:</asp:Label>
                                    </asp:TableCell><asp:TableCell Width="20%" VerticalAlign="Top">
                                            <asp:Label runat="server" forecolor="Red"><%# DataBinder.Eval(Container.DataItem,"ProductDate") %></asp:Label>
                                    </asp:TableCell><asp:TableCell Width="16%" VerticalAlign="Top" Wrap="False">
                                        <asp:Label ID="Label3" runat="server" ForeColor="Gray">Quantity Available:</asp:Label>
                                        &nbsp;<asp:Label ID="Label8" runat="server" ForeColor="Navy"><%# Format(DataBinder.Eval(Container.DataItem,"QtyAvailable"),"#,##0") %></asp:Label>
                                    </asp:TableCell><asp:TableCell Width="5%"></asp:TableCell></asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell VerticalAlign="Top" Wrap="False">
                                        <asp:Label ID="Label13" runat="server" ForeColor="Gray">Description:</asp:Label>
                                    </asp:TableCell><asp:TableCell VerticalAlign="Top" ColumnSpan="3" RowSpan="2">
                                        <asp:Label ID="Label14" runat="server" ForeColor="Navy"><%# DataBinder.Eval(Container.DataItem,"Description") %></asp:Label>
                                    </asp:TableCell><asp:TableCell VerticalAlign="Bottom" Wrap="False" HorizontalAlign="Right"
                                        RowSpan="2">
                                        <asp:Button ID="btn_AddToBasket2" runat="server" Text="add to basket" ToolTip="add product to current order"
                                            OnClick="btn_AddAssocItemToBasket_Click" />
                                    </asp:TableCell><asp:TableCell RowSpan="2"></asp:TableCell></asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell VerticalAlign="Top">
                                        <asp:HyperLink runat="server" ID="hlnk_AssocPDF" Target="_blank" ForeColor="blue"
                                            ToolTip="click here to view a PDF document for this product">view pdf</asp:HyperLink>
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow>
                                    <asp:TableCell></asp:TableCell><asp:TableCell ColumnSpan="6" VerticalAlign="Top">
                                            <hr />
                                    </asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
                            </asp:Table>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </asp:Panel>
    </asp:Panel>
    <asp:Panel ID="pnlRequestAuthorisation" runat="server" Visible="False" Width="100%">
        <asp:Label ID="Label9" runat="server" Font-Size="X-Small" Font-Names="Arial" ForeColor="#0000C0"
            Font-Bold="True">Request Authorisation</asp:Label><br />
        <br />
        <asp:Label ID="Label13" runat="server" Font-Size="XX-Small" Font-Names="Arial">Some items in your basket can only be ordered where authorisation has been granted. Authorisation is required for the products below. Click Request
             Authorisation to start this process, or remove these items from your basket. You will receive an email when authorisation has been granted to order these items, after which you can place your order.</asp:Label><br />
        <br />
        <br />
        <asp:GridView ID="gvRequestAuthorisation" runat="server" AutoGenerateColumns="False"
            Width="95%" Font-Names="Verdana" Font-Size="XX-Small" CellPadding="2">
            <Columns>
                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True"
                    SortExpression="ProductCode" />
                <asp:BoundField DataField="ProductDate" HeaderText="Product Date" ReadOnly="True"
                    SortExpression="ProductDate" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                    SortExpression="ProductDescription" />
                <asp:TemplateField HeaderText="Quantity">
                    <ItemTemplate>
                        <asp:TextBox ID="tbAuthorisationQuantity" Text='<%# Bind("Quantity") %>' runat="server"
                            Width="64px"></asp:TextBox></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Status">
                    <ItemTemplate>
                        <asp:Label ID="lblAuthorisationNotes" Text='<%# Eval("Notes") %>' runat="server"
                            ForeColor="Red"></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Request Authorisation">
                    <ItemTemplate>
                        <asp:CheckBox ID="cbRequestAuthorisation" Checked='<%# Bind("RequestAuthorisation") %>'
                            runat="server" />
                        <asp:HiddenField ID="hidLogisticProductKey" Value='<%# Eval("ProductKey") %>' runat="server" />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
        <br />
        <table>
            <tr>
                <td style="width: 20%; font-size: xx-small; font-family: Verdana;">
                    Note to authoriser (optional):
                </td>
                <td style="width: 181px">
                    <asp:TextBox ID="tbNoteToAuthoriser" runat="server" TextMode="MultiLine" Width="300px"></asp:TextBox>
                </td>
                <td rowspan="2" style="width: 80px">
                    <asp:LinkButton ID="lnkbtnRequestAuthBackToProducts" runat="server" OnClick="lnkbtnRequestAuthBackToProducts_Click"
                        Font-Names="Verdana" Font-Size="XX-Small">back&nbsp;to&nbsp;products</asp:LinkButton><asp:LinkButton
                            ID="lnkbtnRequestAuthShowBasket" runat="server" OnClick="lnkbtnRequestAuthShowBasket_Click"
                            Font-Names="Verdana" Font-Size="XX-Small">view&nbsp;basket</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td style="height: 26px">
                </td>
                <td style="height: 26px; width: 181px;">
                    <asp:Button ID="btnRequestAuthorisation" runat="server" Text="request authorisation"
                        OnClick="btnRequestAuthorisation_Click" />
                </td>
            </tr>
        </table>
        <br />
        &nbsp;<br />
    </asp:Panel>
    <asp:Panel ID="pnlEmptyBasket" runat="server" Visible="False" Width="100%">
        <asp:Table ID="tabEmptyBasket" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Width="100%" CellPadding="2" CellSpacing="1">
            <asp:TableRow runat="server">
                <asp:TableCell Width="5%" runat="server"></asp:TableCell><asp:TableCell Width="45%"
                    runat="server"></asp:TableCell><asp:TableCell Width="45%" runat="server"></asp:TableCell><asp:TableCell
                        Width="5%" runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server"></asp:TableCell><asp:TableCell ColumnSpan="2" runat="server">
                    <asp:Label ID="Label19" runat="server" Font-Size="Small" ForeColor="Navy">Basket</asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="Label20" runat="server" Font-Size="X-Small" ForeColor="Navy">Your basket is empty. Browse available products and add them to your basket.</asp:Label>
                </asp:TableCell><asp:TableCell runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server"></asp:TableCell><asp:TableCell runat="server"></asp:TableCell><asp:TableCell
                    HorizontalAlign="Right" runat="server">
                    <br />
                    <asp:Button ID="Button3" runat="server" OnClick="btn_ContinueWithOrder_click" Text="return to products"
                        ToolTip="go back to browse the stock items" />
                </asp:TableCell><asp:TableCell runat="server"></asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlDeliveryAddress" runat="server" Visible="False" Width="100%">
        <table id="tabDeliveryAddress" width="100%" style="font-family: Verdana; font-size: x-small;
            color: Gray" cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 15%">
                </td>
                <td style="width: 29%">
                </td>
                <td style="width: 2%">
                </td>
                <td style="width: 15%">
                </td>
                <td style="width: 29%">
                </td>
                <td style="width: 5%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <asp:Label ID="Label59" runat="server" Font-Size="Small" ForeColor="Navy">Instructions</asp:Label><br />
                    <br />
                    <asp:Label ID="Label60" runat="server" Font-Size="XX-Small" ForeColor="Navy">Please tell us where to deliver your order. When delivering to a company we need a contact name to prevent the consignment being refused on security grounds.
                        Also, please note that we cannot deliver to a P.O. Box and so a full delivery address is essential.</asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPostCodeLookupStartLine" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPostCodeLookupPreamble" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <asp:Label ID="Label61" runat="server" Font-Size="X-Small">Post Code Lookup</asp:Label><br />
                    <br />
                    <asp:Label ID="Label62" runat="server" Font-Size="XX-Small" ForeColor="Navy" Text="For UK destinations that are not already in your address book, enter the post code, then click <b>find address</b>. Select the required destination from the list of matching addresses." />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPostCodeLookupInput" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label63" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">Post Code:</asp:Label>&nbsp;
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="6" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPostCodeLookup" Width="100%" MaxLength="10"></asp:TextBox>
                </td>
                <td>
                </td>
                <td colspan="2">
                    <asp:Button ID="btnFindAddress" runat="server" Text="find address" CausesValidation="false"
                        OnClick="btnFindAddress_Click" />
                    &nbsp;<asp:Label ID="lblLookupError" runat="server" Visible="False" ForeColor="Red"
                        Font-Size="XX-Small" Font-Names="Verdana"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPostCodeLookupOutput" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label64" runat="server" Text="Select a destination:" Font-Size="XX-Small"
                        Font-Names="Verdana" />
                </td>
                <td colspan="5">
                    <asp:ListBox ID="lbLookupResults" Font-Size="XX-Small" Font-Names="Verdana" BackColor="LightYellow"
                        runat="server" Rows="10" Width="320px" OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged"
                        AutoPostBack="true"></asp:ListBox>
                </td>
            </tr>
            <tr id="trPostCodeLookupFinishLine" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trMultiAddressOrder" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <br />
                    <asp:Label ID="Label113" runat="server" Font-Size="X-Small" Font-Names="Verdana">To send an order to several destinations</asp:Label>&nbsp;<asp:Label
                        ID="Label114" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy"
                        Text=" select a distribution list:"></asp:Label>&nbsp;<asp:DropDownList ID="ddlDistributionList"
                            runat="server" AutoPostBack="true" Font-Names="Verdana" Font-Size="XX-Small"
                            OnSelectedIndexChanged="ddlDistributionList_SelectedIndexChanged" />
                    &nbsp;<asp:Label ID="Label115" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="otherwise enter the delivery address." ForeColor="Navy"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <br />
                    <asp:Label ID="Label65" runat="server" Font-Size="X-Small" Font-Names="Verdana">Delivery Address</asp:Label>
                </td>
                <td colspan="4">
                    <br />
                    <asp:LinkButton ID="lnkbtnGetFromSharedAddressBook" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" OnClick="lnkbtnGetFromSharedAddressBook_Click" CausesValidation="False">view shared address book</asp:LinkButton>&nbsp;
                    &nbsp;<asp:LinkButton ID="lnkbtnGetFromPersonalAddressbook" OnClick="btn_GetFromPersonalAddressBook_click"
                        runat="server" CausesValidation="False" ForeColor="Blue" Font-Size="XX-Small"
                        Font-Names="Verdana">view personal address book</asp:LinkButton>&nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td />
                <td>
                    <asp:Label ID="Label69" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">Country:</asp:Label>&nbsp;
                    <asp:CompareValidator ID="cvCneeCountry" runat="server" ControlToValidate="ddlCneeCountry"
                        Font-Names="Verdana" Font-Size="XX-Small" Operator="NotEqual" ValueToCompare="0">#</asp:CompareValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlCneeCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" Width="90%" AutoPostBack="True" OnSelectedIndexChanged="ddlCneeCountry_SelectedIndexChanged" />
                    &nbsp;<asp:LinkButton ID="lnkbtnCneeCountryUK" runat="server" CausesValidation="False"
                        Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnCneeCountryUK_Click">UK</asp:LinkButton>
                </td>
                <td />
                <td>
                    <asp:Label ID="Label73" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">City:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvCneeCity" runat="server" ControlToValidate="txtCneeCity"
                        Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeCity" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="5" Width="100%"></asp:TextBox>
                </td>
                <td />
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendRegion" runat="server" Font-Names="Verdana" Font-Size="XX-Small">County / Region:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvRegion" runat="server" ControlToValidate="ddlUSStatesCanadianProvinces"
                        Enabled="false" Font-Size="XX-Small" Font-Names="Verdana">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeState" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="6" Width="100%" />
                    <asp:DropDownList ID="ddlUSStatesCanadianProvinces" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Visible="False" TabIndex="6" OnSelectedIndexChanged="ddlUSStatesCanadianProvinces_SelectedIndexChanged" />
                    <asp:Label ID="lblLegendNewYorkCity" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="NEW YORK (NY)" Visible="False" TabIndex="6" />
                </td>
                <td />
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label66" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">Company:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvCneeName" runat="server" ControlToValidate="txtCneeName"
                        Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="1" Width="100%" />
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendPostcodeZipcode" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" ForeColor="Red">Post Code:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPostCodeLookup" runat="server" ControlToValidate="txtCneePostCode"
                        Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="txtCneePostCode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="10" TabIndex="7" Width="100%"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label68" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">Addr 1:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvCneeAddr1" runat="server" ControlToValidate="txtCneeAddr1"
                        Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeAddr1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="2" Width="100%"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label72" runat="server" ForeColor="Red" Font-Size="XX-Small" Font-Names="Verdana"
                        Font-Bold="True">Attn of:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvCneeCtcName" runat="server" ControlToValidate="txtCneeCtcName"
                        Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="8" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCneeCtcName" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label70" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Addr 2:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeAddr2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="3" Width="100%"></asp:TextBox>
                </td>
                <td />
                <td>
                    <asp:Label ID="lblLegendCneeTel" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Contact Tel:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvCneeTel" runat="server" ControlToValidate="txtCneeTel"
                        Font-Size="XX-Small" Enabled="false">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="9" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCneeTel" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td>
                    <asp:Label ID="Label71" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Addr 3:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="txtCneeAddr3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" MaxLength="50" TabIndex="4" Width="100%"></asp:TextBox>
                </td>
                <td />
                <td>
                    <asp:Label ID="Label76" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Contact Email:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="10" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCneeEmail" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td />
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trStandardCheckoutCustRef1" runat="server" visible="true">
                <td />
                <td>
                    <asp:Label ID="lblLegendCheckoutCustomerRef1" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Customer Ref 1:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvCheckoutCustomerRef1" runat="server" Enabled="false" EnableClientScript="false"
                            ControlToValidate="txtCustRef1" Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCustRef1" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td />
                <td>
                    <asp:Label ID="lblLegendCheckoutCustomerRef2" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Customer Ref 2:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvCheckoutCustomerRef2" runat="server" Enabled="false" EnableClientScript="false"
                            ControlToValidate="txtCustRef2" Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCustRef2" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trStandardCheckoutCustRef2" runat="server" visible="true">
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendCheckoutCustomerRef3" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Customer Ref 3:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvCheckoutCustomerRef3" runat="server" Enabled="false" EnableClientScript="false"
                            ControlToValidate="txtCustRef3" Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCustRef3" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendCheckoutCustomerRef4" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Customer Ref 4:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvCheckoutCustomerRef4" runat="server" Enabled="false" EnableClientScript="false"
                            ControlToValidate="txtCustRef4" Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtCustRef4" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration1Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label57" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Cost Centre:</asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration1CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsVar1CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" OnTextChanged="ddlPerCustomerConfiguration1CostCentre_TextChanged">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label58" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">choose from list, or type:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration1CostCentre" runat="server" ControlToValidate="tbPerCustomerConfiguration1CostCentre"
                            Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration1CostCentre" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration2Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label5" runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="red">Booking Ref:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                        ID="rfvPerCustomerConfiguration2BookingRef" runat="server" ControlToValidate="tbPerCustomerConfiguration2BookingRef"
                        Font-Size="XX-Small" Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration2BookingRef" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label31" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref A:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration2AdditionalRefA" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration2Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label32" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref B:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration2AdditionalRefB" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label33" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref C:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration2AdditionalRefC" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration3Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label83" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Cost Centre:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration3CostCentre"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration3CostCentre" Font-Size="XX-Small"
                            Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration3CostCentre" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label84" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref A:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration3AdditionalRefA" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration4Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <asp:Label ID="Label133" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">NOTE: You will choose the delivery service (STANDARD or EXPRESS) at the order confirmation stage.</asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration3Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label85" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref B:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration3AdditionalRefB" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label86" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref C:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration3AdditionalRefC" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration5Checkout1" runat="server" visible="false">
                <td style="height: 24px">
                </td>
                <td style="height: 24px">
                    <asp:Label ID="Label91" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Cost Centre:</asp:Label>
                </td>
                <td style="height: 24px">
                    <asp:DropDownList ID="ddlPerCustomerConfiguration5CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsVar5CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" OnTextChanged="ddlPerCustomerConfiguration5CostCentre_TextChanged">
                    </asp:DropDownList>
                </td>
                <td style="height: 24px">
                </td>
                <td style="height: 24px">
                    <asp:Label ID="Label92" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">choose from list, or type:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration5CostCentre" runat="server" ControlToValidate="tbPerCustomerConfiguration5CostCentre"
                            Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td style="height: 24px">
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration5CostCentre" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td style="height: 24px">
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration6Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label410cimax" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">Department/Market:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration6Department"
                            runat="server" ControlToValidate="ddlPerCustomerConfiguration6Department" Font-Names="Arial"
                            Font-Size="XX-Small" InitialValue="- please select -"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration6Department" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="XdsVar6Departments"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label130cimax" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Cost centre:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration6CostCentre"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration6CostCentre" Font-Names="Verdana"
                            Font-Size="XX-Small">#</asp:RequiredFieldValidator><asp:RegularExpressionValidator
                                ID="revPerCustomerConfiguration6CostCentre" ValidationExpression="(\d\d\d/\d\d\d\d/\d\d\d\d\d)|(\d\d\d/\d\d\d\d/\d\d\d\d\d/\d\d\d\d)"
                                runat="server" ControlToValidate="tbPerCustomerConfiguration6CostCentre" Font-Size="XX-Small"
                                Font-Names="Arial" Enabled="False"> ## {123/1234/12345(/1234)}</asp:RegularExpressionValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration6CostCentre" Width="100%" MaxLength="19"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration6Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label410cimareference" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">Reference*:</asp:Label><asp:RequiredFieldValidator ID="RequiredFieldValidator2"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration6Reference" Font-Names="Verdana"
                            Font-Size="XX-Small"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration6Reference" Width="100%" MaxLength="50" />
                </td>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="Label410cimarefadvice" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">* Please note that this description will appear on the invoice to enable you to authorise it.</asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration7Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label4" runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Red">Service Level:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                        ID="rfvPerCustomerConfiguration7ServiceLevel" ControlToValidate="ddlPerCustomerConfiguration7ServiceLevel"
                        runat="server" Font-Names="Arial" Font-Size="XX-Small" InitialValue="- please select -"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration7ServiceLevel" runat="server" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration7ServiceLevel_SelectedIndexChanged">
                        <asp:ListItem>- please select -</asp:ListItem>
                        <asp:ListItem>ECONOMY</asp:ListItem>
                        <asp:ListItem>EXPRESS</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration18Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label4vsal18" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red">Service Level:</asp:Label><asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration18ServiceLevel"
                            ControlToValidate="ddlPerCustomerConfiguration18ServiceLevel" runat="server"
                            Font-Names="Arial" Font-Size="XX-Small" InitialValue="- please select -"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration18ServiceLevel" runat="server" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration18ServiceLevel_SelectedIndexChanged">
                        <asp:ListItem>- please select -</asp:ListItem>
                        <asp:ListItem>ECONOMY</asp:ListItem>
                        <asp:ListItem>EXPRESS</asp:ListItem>
                        <asp:ListItem>MAIL</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration8Checkout1" runat="server" visible="false">
                <td style="height: 24px">
                </td>
                <td style="height: 24px">
                    <asp:Label ID="Label91c" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Booking Ref:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator3111"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration8BookingRef" Font-Names="Verdana"
                            Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td style="height: 24px">
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration8BookingRef" Width="100%" MaxLength="20"></asp:TextBox>
                </td>
                <td style="height: 24px">
                </td>
                <td style="height: 24px">
                    <asp:Label ID="Label921" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">PCID:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator311"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration8PCID" Font-Names="Verdana"
                            Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td style="height: 24px">
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration8PCID" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td style="height: 24px">
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration8Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label9100" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Rating:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator300"
                            runat="server" ControlToValidate="ddlPerCustomerConfiguration8Rating" InitialValue="- please select -"
                            Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration8Rating" runat="server" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana">
                        <asp:ListItem Selected="true">- please select -</asp:ListItem>
                        <asp:ListItem Selected="False">INVESTOR</asp:ListItem>
                        <asp:ListItem Selected="False">NOSC</asp:ListItem>
                        <asp:ListItem Selected="False">POTENTIAL</asp:ListItem>
                        <asp:ListItem Selected="False">RO</asp:ListItem>
                        <asp:ListItem Selected="False">SIGNED</asp:ListItem>
                        <asp:ListItem Selected="False">OTHER</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label920" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">RO:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator31"
                            runat="server" ControlToValidate="ddlPerCustomerConfiguration8RO" InitialValue="- please select -"
                            Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration8RO" runat="server" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana">
                        <asp:ListItem Selected="true">- please select -</asp:ListItem>
                        <asp:ListItem Selected="False">AUGE</asp:ListItem>
                        <asp:ListItem Selected="False">Chicago</asp:ListItem>
                        <asp:ListItem Selected="False">Compulsory</asp:ListItem>
                        <asp:ListItem Selected="False">GRS</asp:ListItem>
                        <asp:ListItem Selected="False">Hong Kong</asp:ListItem>
                        <asp:ListItem Selected="False">London </asp:ListItem>
                        <asp:ListItem Selected="False">Middle East</asp:ListItem>
                        <asp:ListItem Selected="False">Montevideo</asp:ListItem>
                        <asp:ListItem Selected="False">None</asp:ListItem>
                        <asp:ListItem Selected="False">Swiss Sales</asp:ListItem>
                        <asp:ListItem Selected="False">Tokyo</asp:ListItem>
                        <asp:ListItem Selected="False">TSS</asp:ListItem>
                        <asp:ListItem Selected="False">Other</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration8Checkout3" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label910" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">MDS Order Ref:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="RequiredFieldValidator30"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration8MDSOrderRef" Font-Names="Verdana"
                            Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration8MDSOrderRef" Width="100%" MaxLength="20"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration12Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label134" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Cost Code:</asp:Label>&nbsp;<asp:RequiredFieldValidator ID="rfdPerCustomerConfiguration12CostCentre"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration12CostCentre" Font-Size="XX-Small"
                            Font-Names="Arial"> #</asp:RequiredFieldValidator><asp:RegularExpressionValidator
                                ID="revPerCustomerConfiguration12CostCentre" ValidationExpression="(\w\w\d\d\d\d\.\d\d\d\d)|(\d\d\d\d\d\d\d.\d\d\d\d\.\d\d\d)"
                                runat="server" ControlToValidate="tbPerCustomerConfiguration12CostCentre" Font-Size="XX-Small"
                                Font-Names="Arial"> # (wrong format)</asp:RegularExpressionValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration12CostCentre" Width="90%" MaxLength="50"></asp:TextBox><a
                            runat="server" id="aPerCustomerConfiguration12HelpCostCode" visible="true" onmouseover="return escape('Valid cost codes have the format:<br /><br /> &lt; letter &gt;&lt; letter &gt; &lt; digit &gt;&lt; digit &gt; &lt; digit &gt;&lt; digit &gt; <b>.</b> &lt; digit &gt;&lt; digit &gt; &lt; digit &gt;&lt; digit &gt; <br /><br />OR<br /><br />&lt; digit &gt;&lt; digit &gt;&lt; digit &gt;&lt; digit &gt;&lt; digit &gt; &lt; digit &gt;&lt; digit &gt; <b>.</b> &lt; digit &gt;&lt; digit &gt; &lt; digit &gt;&lt; digit &gt; <b>.</b> &lt; digit &gt;&lt; digit &gt; &lt; digit &gt;')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label135" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref A:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration12AdditionalRefA" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration12Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label136" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref B:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration12AdditionalRefB" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="Label137" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Additional Customer Ref C:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration12AdditionalRefC" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration17Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label8117" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">Service Level:</asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration17ServiceLevel" runat="server" Font-Size="XX-Small"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration17ServiceLevel_SelectedIndexChanged">
                        <asp:ListItem Value="courier">Standard Shipping (Courier)</asp:ListItem>
                        <asp:ListItem Value="mail">Mail</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration18Checkout1" runat="server" visible="false">
                <td />
                <td colspan="5">
                    <asp:Label ID="Label9699" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">NOTE: Your consignment will be sent ECONOMY unless you specify otherwise in Special Instructions</asp:Label>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration18Checkout19" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label5719" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Cost Centre:</asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration19CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsVar19CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" OnTextChanged="ddlPerCustomerConfiguration19CostCentre_TextChanged">
                    </asp:DropDownList>
                </td>
                <td />
                <td>
                    <asp:Label ID="Label5819" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">choose from list, or type:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration19CostCentre" Enabled="false" runat="server" ControlToValidate="tbPerCustomerConfiguration19CostCentre"
                            Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration19CostCentre" Width="100%" MaxLength="25" />
                </td>
                <td />
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration20Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label5720" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Cost Centre:</asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration20CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" Font-Size="XX-Small" DataSourceID="xdsVar20CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana">
                    </asp:DropDownList>
                    &nbsp;<asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration20CostCentre" runat="server"
                        ControlToValidate="ddlPerCustomerConfiguration20CostCentre" ErrorMessage="###"
                        Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -"></asp:RequiredFieldValidator>
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration21Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="lblAPerCustomerConfiguration21CostCentre" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="red">Cost Centre:</asp:Label>
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration21CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" Font-Size="XX-Small" DataSourceID="xdsVar21CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana">
                    </asp:DropDownList>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration21CostCentre" runat="server"
                        ControlToValidate="ddlPerCustomerConfiguration21CostCentre" ErrorMessage="###"
                        Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td />
                <td>
                    <asp:Label ID="lblBPerCustomerConfiguration21CostCentre" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana">Customer Ref:</asp:Label>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration21CustRef" Width="100%" MaxLength="50" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration22Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label134rio" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Cost Centre:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration22CostCentre" runat="server"
                        Enabled="true" ControlToValidate="tbPerCustomerConfiguration22CostCentre" Font-Size="XX-Small"
                        Font-Names="Arial"> #</asp:RequiredFieldValidator><asp:RegularExpressionValidator
                            ID="revPerCustomerConfiguration22RequestedBy" Enabled="true" ValidationExpression="^\d\d\d\d\d\d\d\d(\d)?(\d)?$"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration22CostCentre" Font-Size="XX-Small"
                            Font-Names="Arial"> # (8 - 10 digits)</asp:RegularExpressionValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration22CostCentre" Width="90%" MaxLength="13" />
                    <a runat="server" id="aPerCustomerConfiguration22HelpCostCentre" visible="true" onmouseover="return escape('Valid Cost Centre codes are 8 - 10 digits')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td />
                <td>
                    <asp:Label ID="Label135rio" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Requested By:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration22RequestedBy" runat="server"
                        ControlToValidate="tbPerCustomerConfiguration22RequestedBy" Font-Size="XX-Small"
                        Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="tbPerCustomerConfiguration22RequestedBy" runat="server" ForeColor="Navy"
                        TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana" Width="100%" MaxLength="50" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration23Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label134arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Font-Bold="true" ForeColor="red">Cost Centre:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration23CostCentre" runat="server" ControlToValidate="ddlPerCustomerConfiguration23CostCentre"
                            ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration23CostCentre" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsVar23CostCentres"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" />
                </td>
                <td />
                <td>
                    <asp:Label ID="Label135arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Font-Bold="true" ForeColor="red">Category:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration23Category" runat="server" ControlToValidate="ddlPerCustomerConfiguration23Category"
                            ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration23Category" runat="server" DataTextField="name"
                        DataValueField="name" AutoPostBack="true" Font-Size="XX-Small" DataSourceID="xdsVar23Categories"
                        ForeColor="Navy" TabIndex="12" Font-Names="Verdana" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration23Checkout2" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label136arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Font-Bold="true" ForeColor="red">PO Number:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfdPerCustomerConfiguration23PONumber" runat="server" ControlToValidate="tbPerCustomerConfiguration23PONumber"
                            Font-Names="Verdana" Font-Size="XX-Small"> ###</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration23PONumber" Width="100%" MaxLength="25" />
                </td>
                <td />
                <td>
                </td>
                <td>
                    <asp:Label ID="Label137arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Type any further order details in the Special Instructions box</asp:Label>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration24Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label134pvx" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Reference:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration24Reference" runat="server"
                        Enabled="true" ControlToValidate="tbPerCustomerConfiguration24Reference" Font-Size="XX-Small"
                        Font-Names="Arial"> #</asp:RequiredFieldValidator><asp:RegularExpressionValidator
                            ID="revPerCustomerConfiguration24Reference" Enabled="true" ValidationExpression="^[a-zA-Z][a-zA-Z]\d\d\d\d$"
                            runat="server" ControlToValidate="tbPerCustomerConfiguration24Reference" Font-Size="XX-Small"
                            Font-Names="Arial"> # (2 letters, 4 digits)</asp:RegularExpressionValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration24Reference" Width="90%" MaxLength="6" />
                    <a runat="server" id="aPerCustomerConfiguration24HelpReference" visible="true" onmouseover="return escape('Valid References are 2 letters followed by 4 digits')"
                        style="color: gray; cursor: help">&nbsp;?&nbsp;</a>
                </td>
                <td />
                <td>
                    <asp:Label ID="Label135pvx" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Requested By:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration24RequestedBy" runat="server"
                        ControlToValidate="tbPerCustomerConfiguration24RequestedBy" Font-Size="XX-Small"
                        Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="tbPerCustomerConfiguration24RequestedBy" runat="server" ForeColor="Navy"
                        TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana" Width="100%" MaxLength="25" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration24Checkout2" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="Label134pv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Recipient Name:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration24RecipientName" runat="server"
                        Enabled="true" ControlToValidate="tbPerCustomerConfiguration24RecipientName"
                        Font-Size="XX-Small" Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration24RecipientName" Width="90%" MaxLength="50" />
                </td>
                <td />
                <td>
                    <asp:Label ID="Label135pv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="red">Products:</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration24Products" runat="server"
                        ControlToValidate="tbPerCustomerConfiguration24Products" Font-Size="XX-Small"
                        Font-Names="Arial"> #</asp:RequiredFieldValidator>
                </td>
                <td>
                    <asp:TextBox ID="tbPerCustomerConfiguration24Products" runat="server" ForeColor="Navy"
                        TabIndex="13" Font-Size="XX-Small" Font-Names="Verdana" Width="100%" MaxLength="50" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration26Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="lblRamblersChooseGroup" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red" Font-Bold="True">Area/Group:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration26RamblersAreaGroup" runat="server" ControlToValidate="ddlRamblersAreaGroup"
                            ErrorMessage="###" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"
                            InitialValue="- please select -" />
                </td>
                <td colspan="4">
                    <asp:LinkButton ID="lnkbtnRamblersGroupAB" CommandArgument="A-B" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">A-B</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupCD" CommandArgument="C-D" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">C-D</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupEG" CommandArgument="E-G" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">E-G</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupHJ" CommandArgument="H-J" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">H-J</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupKL" CommandArgument="K-L" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">K-L</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupMN" CommandArgument="M-N" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">M-N</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupOR" CommandArgument="O-R" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">O-R</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupS" CommandArgument="S" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">S</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupTZ" CommandArgument="T-Z" runat="server" OnClick="lnkbtnRamblersGroup_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" CausesValidation="False">T-Z</asp:LinkButton>&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnRamblersGroupOther" CommandArgument="Other" runat="server"
                        OnClick="lnkbtnRamblersGroup_Click" Font-Names="Verdana" Font-Size="XX-Small"
                        CausesValidation="False">Other</asp:LinkButton>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration26Checkout2" runat="server" visible="false">
                <td />
                <td>
                </td>
                <td>
                    <asp:DropDownList ID="ddlRamblersAreaGroup" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        DataSourceID="xdsVar26RamblersAreaGroups" DataTextField="name">
                        <asp:ListItem>- Select from the alphabetic choices above to find your area/group -</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp;
                </td>
                <td />
                <td>
                </td>
                <td>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration29Checkout1" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="lblIrwinMitchellChooseBudgetCode" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Font-Bold="true" ForeColor="red">Budget Code:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration29IrwinMitchellBudgetCode" runat="server" ControlToValidate="ddlPerCustomerConfiguration29IrwinMitchellBudgetCode"
                            ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration29IrwinMitchellBudgetCode" runat="server"
                        DataTextField="name" DataValueField="name" AutoPostBack="true" Font-Size="XX-Small"
                        DataSourceID="xdsVar29IrwinMitchellBudgetCodes" ForeColor="Navy" TabIndex="12"
                        Font-Names="Verdana" />
                </td>
                <td />
                <td>
                    <asp:Label ID="lblIrwinMitchellChooseDepartment" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Font-Bold="true" ForeColor="red">Department:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration29Department" runat="server" ControlToValidate="ddlPerCustomerConfiguration29IrwinMitchellDepartment"
                            ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration29IrwinMitchellDepartment" runat="server"
                        DataTextField="name" DataValueField="name" AutoPostBack="true" Font-Size="XX-Small"
                        DataSourceID="xdsVar29IrwinMitchellDepartments" ForeColor="Navy" TabIndex="12"
                        Font-Names="Verdana" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration29Checkout2" runat="server" visible="false">
                <td />
                <td>
                    <asp:Label ID="lblLegendIrwinMitchellChooseBudgetCodeEvent" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Font-Bold="true" ForeColor="red">Event:</asp:Label>&nbsp;<asp:RequiredFieldValidator
                            ID="rfvPerCustomerConfiguration29Event" runat="server" ControlToValidate="tbPerCustomerConfiguration29Event"
                            ErrorMessage="###" Font-Size="XX-Small" Font-Names="Verdana" />
                </td>
                <td>
                    <asp:TextBox runat="server" ForeColor="Navy" TabIndex="12" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="tbPerCustomerConfiguration29Event" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td />
                <td>
                    <asp:Label ID="lblIrwinMitchellBDName" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Font-Bold="true">BD Name:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbPerCustomerConfiguration29BDName" runat="server" 
                        Font-Size="XX-Small" Font-Names="Verdana" Width="100%" MaxLength="50" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration30Checkout1" runat="server" visible="false">
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendJupiter" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        Font-Bold="true" ForeColor="red">Print Service Level</asp:Label>&nbsp;
                    <asp:RequiredFieldValidator ID="rfvPerCustomerConfiguration30PrintServiceLevel" runat="server"
                        ErrorMessage="###" Font-Size="XX-Small" Font-Names="Verdana" ControlToValidate="ddlPerCustomerConfiguration30PrintServiceLevel"
                        InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration30PrintServiceLevel" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration30PrintServiceLevel_SelectedIndexChanged">
                    </asp:DropDownList>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration30BudgetCode" runat="server" Font-Bold="true"
                        Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red">Budget Code:</asp:Label><asp:RequiredFieldValidator
                            ID="RequiredFieldValidator30BudgetCode" runat="server" ControlToValidate="ddlPerCustomerConfiguration30JupiterBudgetCode"
                            ErrorMessage="###" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="- please select -" />
                </td>
                <td>
                    <asp:DropDownList ID="ddlPerCustomerConfiguration30JupiterBudgetCode" runat="server"
                        DataTextField="name" DataValueField="name" AutoPostBack="true" Font-Size="XX-Small"
                        DataSourceID="xdsVar30JupiterBudgetCode" ForeColor="Navy" TabIndex="12" Font-Names="Verdana" />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration30Checkout2" runat="server" visible="false">
                <td />
                <td>
                </td>
                <td>
                    <asp:Label ID="Label137arth0" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red">To calculate total elapsed time from ordering to delivery, add the Print Service Level time to the transit time to your destination, as listed in the...</asp:Label><br />
                    <asp:HyperLink ID="HyperLink4" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" NavigateUrl="http://my.transworld.eu.com/jupiter/customer_furniture/jupiter/globaltransitguide.pdf"
                        Target="_blank">GLOBAL TRANSIT GUIDE</asp:HyperLink>
                </td>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendPerCustomerConfiguration30TotalPrintCost2" runat="server"
                        Font-Names="Verdana" Font-Size="XX-Small">Total Print Cost:</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration30TotalPrintCost" runat="server" Font-Bold="true"
                        Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration1Checkout2" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <span style="font-size: xx-small; color: red; font-family: Verdana">Please click <a
                        href="javascript:void(window.open('http://my.transworld.eu.com/info/transit_times.pdf','','resizable=yes,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,fullscreen=no,dependent=no'))">
                        <strong>here</strong></a> to view standard delivery times to worldwide locations.
                        Note that all shipments will be sent using standard delivery times unless advised
                        otherwise. If your shipment requires urgent attention, or is a low priority shipment
                        that can go via alternative lower cost transportation, please notify in "Special
                        Instructions" below.</span><br />
                    <br />
                    <asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="Required Delivery Date (optional): " />
                    <a href="javascript:;" onclick="window.open('PopupCalendar4.aspx?textbox=txtSpecialInstructions&mode=append&prefixtext=REQUIRED%20DELIVERY%20DATE:%20','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img1" alt="" style="border: none" src="~/images/SmallCalendar.gif" runat="server"
                            ie:visible="true" visible="false" /></a> <a onmouseover="return escape('Use the Required Delivery Date feature to insert the preferred delivery date into the Special Instructions box.Click the calendar icon to display the calendar, then select a date.')"
                                style="color: gray; cursor: help">&nbsp;?&nbsp;</a> &nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr id="trPerCustomerConfiguration0CheckoutDeliveryDateCalendar" runat="server" visible="false">
                <td>
                </td>
                <td colspan="5">
                    <span style="font-size: xx-small; color: red; font-family: Verdana">Please click <a
                        href="javascript:void(window.open('http://my.transworld.eu.com/info/transit_times.pdf','','resizable=yes,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,fullscreen=no,dependent=no'))">
                        <strong>here</strong></a> to view standard delivery times to worldwide locations.
                        All shipments will be sent using standard delivery times unless advised otherwise,
                        e.g. in "Special Instructions" below.</span><br />
                    <br />
                    <span id="spnPerCustomerConfiguration0CheckoutDeliveryDateCalendar" runat="server"
                        visible="true">
                        <asp:Label ID="lblPerCustomerConfiguration0CheckoutDeliveryDateCalendar" runat="server"
                            Font-Names="Verdana" Font-Size="XX-Small" Text="Required Delivery Date (optional): " />
                        <a id="aPerCustomerConfiguration0CheckoutDeliveryDateCalendar01" href="javascript:;"
                            onclick="window.open('PopupCalendar4.aspx?textbox=txtSpecialInstructions&mode=append&prefixtext=REQUIRED%20DELIVERY%20DATE:%20','cal','width=300,height=305,left=270,top=180')">
                            <img id="imgPerCustomerConfiguration0CheckoutDeliveryDateCalendar" alt="" style="border: none"
                                src="~/images/SmallCalendar.gif" runat="server" ie:visible="true" visible="false" /></a>
                        <a id="PerCustomerConfiguration0CheckoutDeliveryDateCalendar02" onmouseover="return escape('Use the Required Delivery Date feature to insert the preferred delivery date into the Special Instructions box.Click the calendar icon to display the calendar, then select a date.<br /><br />If you require express courier(1 - 2 days) delivery click the <b>express courier</b> option button. Express courier orders are automatically sent for authorisation.')"
                            style="color: gray; cursor: help">&nbsp;?&nbsp;</a> </span>&nbsp;<asp:RadioButton
                                ID="rblPerCustomerConfiguration0CheckoutNextDayDelivery00" runat="server" Font-Names="Verdana"
                                Font-Size="XX-Small" Text="standard delivery" GroupName="rblPerCustomerConfiguration0CheckoutNextDayDelivery"
                                AutoPostBack="True" Checked="True" OnCheckedChanged="rblPerCustomerConfiguration0CheckoutNextDayDelivery00_CheckedChanged" />
                    <asp:RadioButton ID="rblPerCustomerConfiguration0CheckoutNextDayDelivery01" runat="server"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="express courier (1 - 2 days) delivery &lt;font color=&quot;red&quot;&gt;(your order will be sent for authorisation if you select this option)&lt;/font&gt;"
                        GroupName="rblPerCustomerConfiguration0CheckoutNextDayDelivery" AutoPostBack="True"
                        OnCheckedChanged="rblPerCustomerConfiguration0CheckoutNextDayDelivery01_CheckedChanged" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Label ID="lblLegendSpecialInstructionsCheckout" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" Text="Special Instructions:" />
                    &nbsp;<asp:RequiredFieldValidator ID="rfvInstructions" runat="server" ControlToValidate="txtSpecialInstructions"
                        Enabled="false" Font-Names="Verdana" Font-Size="XX-Small">#</asp:RequiredFieldValidator>
                </td>
                <td colspan="4">
                    <asp:TextBox runat="server" ForeColor="Red" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtSpecialInstructions" Width="100%" MaxLength="1000"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr id="trCheckoutPackingNoteText" runat="server" visible="true">
                <td>
                </td>
                <td>
                    <asp:Label ID="Label82" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Packing Note Text:</asp:Label>
                </td>
                <td colspan="4">
                    <asp:TextBox runat="server" ForeColor="Red" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana"
                        ID="txtShippingInfo" Width="100%" MaxLength="1000"></asp:TextBox>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblCheckOutMessage" runat="server" BackColor="#F9D938" ForeColor="Navy"
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                </td>
                <td colspan="3" align="right">
                    <asp:Button ID="btnConfirmOrder" OnClick="btnConfirmOrder_click" runat="server" ToolTip="click here to complete your order"
                        Text="final check"></asp:Button>
                </td>
                <td>
                </td>
            </tr>
        </table>
        <asp:XmlDataSource ID="xdsVar1CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/BlackRockCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar5CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/KodakCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar6CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/CimaCostCentres/costCentre" />
        <asp:XmlDataSource ID="XdsVar6Departments" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/CimaDepartments/department" />
        <asp:XmlDataSource ID="xdsVar19CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/InsightCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar20CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/DATCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar21CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/UNICRDCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar23CostCentres" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/ArthritisCostCentres/costCentre" />
        <asp:XmlDataSource ID="xdsVar23Categories" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/ArthritisCategories/category" />
        <asp:XmlDataSource ID="xdsVar26RamblersAreaGroups" runat="server" DataFile="~/on_line_picks_config_ramblers.xml"
            XPath="RamblersAreaGroups/RamblersA-B/areaGroup" />
        <asp:XmlDataSource ID="xdsVar29IrwinMitchellBudgetCodes" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/IrwinMitchellBudgetCodes/category" />
        <asp:XmlDataSource ID="xdsVar29IrwinMitchellDepartments" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/IrwinMitchellDepartments/category" />
        <asp:XmlDataSource ID="xdsVar30JupiterBudgetCode" runat="server" DataFile="~/on_line_picks_config.xml"
            XPath="OnLinePicksConfig/JupiterBudgetCodes/category" />
    </asp:Panel>
    <asp:Panel ID="pnlDistributionList" runat="server" Visible="False" Width="100%">
        <asp:Label ID="Label36" runat="server" Font-Size="X-Small" Font-Names="Arial" ForeColor="#0000C0"
            Font-Bold="True">Multiple Delivery Addresses</asp:Label><br />
        <br />
        <asp:Label ID="Label10" runat="server" Font-Size="XX-Small" Font-Names="Arial">For each addressee and stock item, select the quantity to be sent. Addressee rows where the quantity for every item is 0 or blank will be ignored.</asp:Label><br />
        <br />
        <asp:GridView ID="gvDistributionList" AutoGenerateColumns="False" runat="server"
            Font-Names="Verdana" Width="95%" Font-Size="XX-Small" CellPadding="2" OnDataBound="gvDistributionList_DataBound">
            <Columns>
                <asp:TemplateField HeaderText="Item1" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem1" Text='<%# gvDistributionListSetQty(Container.DataItem,1) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item2" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem2" Text='<%# gvDistributionListSetQty(Container.DataItem,2) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item3" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem3" Text='<%# gvDistributionListSetQty(Container.DataItem,3) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item4" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem4" Text='<%# gvDistributionListSetQty(Container.DataItem,4) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item5" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem5" Text='<%# gvDistributionListSetQty(Container.DataItem,5) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item6" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem6" Text='<%# gvDistributionListSetQty(Container.DataItem,6) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item7" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem7" Text='<%# gvDistributionListSetQty(Container.DataItem,7) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item8" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem8" Text='<%# gvDistributionListSetQty(Container.DataItem,8) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item9" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem9" Text='<%# gvDistributionListSetQty(Container.DataItem,9) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Item10" Visible="False">
                    <ItemTemplate>
                        Qty:
                        <asp:TextBox ID="tbItem10" Text='<%# gvDistributionListSetQty(Container.DataItem,10) %>'
                            Font-Names="Arial" Font-Size="XX-Small" runat="server" Width="40px" MaxLength="5"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Service">
                    <ItemTemplate>
                        <asp:DropDownList ID="ddlServiceLevel" Font-Names="Arial" Font-Size="XX-Small" runat="server">
                            <asp:ListItem>STANDARD</asp:ListItem>
                            <asp:ListItem>PREMIUM</asp:ListItem>
                        </asp:DropDownList>
                    </ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Cost Centre">
                    <ItemTemplate>
                        <asp:TextBox ID="tbDistributionListCostCentre" ForeColor="red" Text='<%# gvDistributionListSetCostCentre(Container.DataItem) %>'
                            runat="server" Font-Names="Arial" Font-Size="XX-Small" MaxLength="40" Width="70px"></asp:TextBox></ItemTemplate>
                    <ItemStyle HorizontalAlign="Center" />
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Addressee">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnAddInfo" runat="server" OnClick="lnkbtnAddInfo_Click">add info</asp:LinkButton><asp:Image
                            ID="imgTransparentSphere" ImageUrl=".\images\transparentsphere11x11.gif" runat="server" />
                        <asp:Image ID="imgRedSphere" ImageUrl=".\images\redsphere11x11.gif" runat="server"
                            Visible="false" />
                        <asp:Label ID="lblCneeName" Text='<%# Eval("Company") %>' runat="server" Font-Bold="true" />&nbsp;<asp:Label
                            ID="lblCneeAddr1" Text='<%# Eval("Addr1") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeAddr2" Value='<%# Eval("Addr2") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeAddr3" Value='<%# Eval("Addr3") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeTown" Value='<%# Eval("Town") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeCounty" Value='<%# Eval("State") %>' runat="server" />
                        <asp:HiddenField ID="hidCneePostCode" Value='<%# Eval("PostCode") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeCountryCode" Value='<%# Eval("CountryKey") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeCtcName" Value='<%# Eval("AttnOf") %>' runat="server" />
                        <asp:HiddenField ID="hidCneePhone" Value='<%# Eval("Telephone") %>' runat="server" />
                        <asp:HiddenField ID="hidCneeEmail" Value='<%# Eval("Email") %>' runat="server" />
                        <asp:HiddenField ID="hidCustomerReference" runat="server" />
                        <asp:HiddenField ID="hidSpecialInstructions" runat="server" />
                        <asp:HiddenField ID="hidPackingNote" runat="server" />
                    </ItemTemplate>
                    <HeaderTemplate>
                        <strong>Addressee</strong> &nbsp; &nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnAddAddressee" runat="server" Font-Bold="false" Font-Names="Arial"
                            Visible="false">[add addressee]</asp:LinkButton></HeaderTemplate>
                </asp:TemplateField>
            </Columns>
            <AlternatingRowStyle BackColor="WhiteSmoke" />
            <EmptyDataTemplate>
                no addresses found</EmptyDataTemplate>
        </asp:GridView>
        <br />
        <span style="font-size: xx-small; font-family: Verdana;">Set all Qty to:
            <asp:TextBox runat="server" Width="50px" ID="tbSetAllQuantities" Font-Size="XX-Small"
                MaxLength="5"></asp:TextBox><asp:Button ID="btnSetAllQuantities" runat="server" Text="go"
                    OnClick="btnSetAllQuantities_Click" /></span> &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkbtnClearAllQuantities"
                        runat="server" OnClick="lnkbtnClearAllQuantities_Click" Font-Names="Verdana"
                        Font-Size="XX-Small">clear all qty</asp:LinkButton>&nbsp; &nbsp;
        &nbsp; &nbsp; <span style="font-size: xx-small; font-family: Verdana;">Set all Service
            to:
            <asp:DropDownList ID="ddlSetAllService" runat="server" AutoPostBack="True" Font-Names="Verdana"
                Font-Size="XX-Small" OnSelectedIndexChanged="ddlSetAllService_SelectedIndexChanged">
                <asp:ListItem>- please select -</asp:ListItem>
                <asp:ListItem>STANDARD</asp:ListItem>
                <asp:ListItem>PREMIUM</asp:ListItem>
            </asp:DropDownList>
        </span>&nbsp; &nbsp;&nbsp; <span style="font-size: xx-small; font-family: Verdana;">
            Set all Cost Centre to:
            <asp:TextBox runat="server" Width="50px" ID="tbSetAllCostCentres" Font-Size="XX-Small"
                MaxLength="5"></asp:TextBox><asp:Button ID="btnSetAllCostCentres" runat="server"
                    Text="go" OnClick="btnSetAllCostCentres_Click" /></span> &nbsp;&nbsp;&nbsp;<asp:LinkButton
                        ID="lnkbtnClearAllCostCentres" runat="server" OnClick="lnkbtnClearAllCostCentres_Click"
                        Font-Names="Verdana" Font-Size="XX-Small">clear all cost centre</asp:LinkButton><asp:Panel
                            ID="pnlUpdateOrder" runat="server" Visible="false" Width="100%" BackColor="#ffff99">
                            <asp:Table ID="tabMultipleDeliveryAddresses" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                                Width="95%" ForeColor="Gray">
                                <asp:TableRow ID="TableRow2" runat="server">
                                    <asp:TableCell ID="TableCell62" Width="15%" runat="server"></asp:TableCell><asp:TableCell
                                        ID="TableCell63" Width="29%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell64"
                                            Width="2%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell65" Width="15%"
                                                runat="server"></asp:TableCell><asp:TableCell ID="TableCell66" Width="29%" runat="server"></asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow3" runat="server">
                                    <asp:TableCell ID="TableCell67" ColumnSpan="5" runat="server">
                                    <hr />
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow11" runat="server">
                                    <asp:TableCell ID="TableCell79q" Wrap="False" runat="server">
                                        <asp:Label ID="lblLegendCustomerRefMultiAddress" runat="server" Font-Size="XX-Small"
                                            Font-Names="Verdana">Customer Ref:</asp:Label>
                                    </asp:TableCell><asp:TableCell ID="TableCell80q" ColumnSpan="4" runat="server">
                                        <asp:TextBox runat="server" ForeColor="Red" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana"
                                            ID="tbMultipleAddressOrderCustomerRef" Width="50%" MaxLength="50"></asp:TextBox>
                                        <asp:CheckBox ID="cbUseCustomerRefForAllDestinations" Font-Size="XX-Small" Font-Names="Verdana"
                                            Text="Use for all destinations" runat="server" />
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow7" runat="server">
                                    <asp:TableCell ID="TableCell79" Wrap="False" runat="server">
                                        <asp:Label ID="lblLegendSpecialInstructionsMultiAddress" runat="server" Font-Size="XX-Small"
                                            Font-Names="Verdana">Special Instructions:</asp:Label>
                                    </asp:TableCell><asp:TableCell ID="TableCell80" ColumnSpan="4" runat="server">
                                        <asp:TextBox runat="server" ForeColor="Red" TabIndex="16" Font-Size="XX-Small" Font-Names="Verdana"
                                            ID="tbMultipleAddressOrderSpecialInstructions" Width="75%" MaxLength="1000"></asp:TextBox>
                                        <asp:CheckBox ID="cbUseSpecialInstructionsForAllDestinations" Font-Size="XX-Small"
                                            Font-Names="Verdana" Text="Use for all destinations" runat="server" />
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow8" runat="server">
                                    <asp:TableCell ID="TableCell81" Wrap="False" runat="server">
                                        <asp:Label ID="Label103" runat="server" Font-Size="XX-Small" Font-Names="Verdana">Packing Note Text:</asp:Label>
                                    </asp:TableCell><asp:TableCell ID="TableCell82" ColumnSpan="4" runat="server">
                                        <asp:TextBox runat="server" TabIndex="17" Font-Size="XX-Small" Font-Names="Verdana"
                                            ID="tbMultipleAddressOrderShippingInfo" Width="75%" MaxLength="1000"></asp:TextBox>
                                        <asp:CheckBox ID="cbUseShippingInfoForAllDestinations" Font-Size="XX-Small" Font-Names="Verdana"
                                            Text="Use for all destinations" runat="server" />
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow5" runat="server">
                                    <asp:TableCell ID="TableCell84qq" runat="server"></asp:TableCell><asp:TableCell ID="TableCell85qq"
                                        runat="server">
                                    </asp:TableCell><asp:TableCell ID="TableCell86qq" runat="server"></asp:TableCell><asp:TableCell
                                        ID="TableCell87qq" HorizontalAlign="Right" runat="server">
                                    </asp:TableCell><asp:TableCell HorizontalAlign="right" ID="TableCell68" runat="server">
                                        <asp:Button ID="btnMultipleAddressOrderUpdateOrder" runat="server" ToolTip="" Text="update order"
                                            OnClick="btnMultipleAddressOrderUpdateOrder_Click"></asp:Button>
                                    </asp:TableCell></asp:TableRow>
                                <asp:TableRow ID="TableRow9" runat="server">
                                    <asp:TableCell ID="TableCell83" ColumnSpan="5" runat="server">
                                    <hr />
                                    </asp:TableCell></asp:TableRow>
                            </asp:Table>
                            <asp:HiddenField ID="hidEditRow" runat="server" />
                        </asp:Panel>
        <asp:Table ID="tabMultipleDeliveryAddresses2" runat="server" Font-Size="X-Small"
            Font-Names="Verdana" Width="95%" ForeColor="Gray">
            <asp:TableRow ID="TableRow4" runat="server">
                <asp:TableCell ID="TableCell622" Width="15%" runat="server"></asp:TableCell><asp:TableCell
                    ID="TableCell632" Width="29%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell642"
                        Width="2%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell625" Width="15%"
                            runat="server"></asp:TableCell><asp:TableCell ID="TableCell626" Width="29%" runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow ID="TableRow10" runat="server">
                <asp:TableCell ID="TableCell84" runat="server"></asp:TableCell><asp:TableCell ID="TableCell85"
                    runat="server">
                    <asp:Label ID="lblMultipleAddressOrderCheckOutMessage" runat="server" BackColor="#F9D938"
                        ForeColor="Navy" Font-Names="Verdana" Font-Size="X-Small"></asp:Label>
                </asp:TableCell><asp:TableCell ID="TableCell86" runat="server"></asp:TableCell><asp:TableCell
                    ID="TableCell87" HorizontalAlign="Right" runat="server">
                </asp:TableCell><asp:TableCell ID="TableCell88" HorizontalAlign="Right" runat="server">
                    <asp:Button ID="btnMultipleAddressOrderFinalConfirmation" runat="server" ToolTip="click here to confirm your order"
                        Text="final confirmation" OnClick="btnMultipleAddressOrderFinalConfirmation_Click">
                    </asp:Button>
                </asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlSearchAddress" runat="server" Visible="False" Width="100%">
        <table id="tabSearchAddressBook" style="width: 100%; font-family: Verdana; font-size: xx-small"
            cellpadding="2" cellspacing="1">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 70%">
                </td>
                <td style="width: 20%">
                </td>
                <td style="width: 5px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblLegendAddressBookType" runat="server" Font-Size="Small" ForeColor="Navy"
                        Text="Address Book" />
                    &nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnUseSharedAddressbook" runat="server" OnClick="lnkbtnUseSharedAddressbook_Click">use shared address book</asp:LinkButton><asp:LinkButton
                        ID="lnkbtnUsePersonalAddressBook" runat="server" OnClick="lnkbtnUsePersonalAddressBook_Click">use personal address book</asp:LinkButton><br />
                    <br />
                    <asp:Label ID="Label99" runat="server" Font-Size="XX-Small" ForeColor="Navy">Use your address book to save re-typing the delivery address.</asp:Label>
                </td>
                <td style="width: 5px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnShowAllAddresses" OnClick="btn_ShowAllAddresses_click" runat="server"
                        ToolTip="click here to show all addresses" Text="show all"></asp:Button>
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label98" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy">Search for:</asp:Label><asp:TextBox runat="server"
                            Width="90px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtSearchCriteriaAddress"
                            ToolTip="search criteria" MaxLength="50"></asp:TextBox><asp:Label ID="Label1x" runat="server"
                                Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy">in</asp:Label><asp:DropDownList
                                    ID="ddlAddressFields" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                    <asp:ListItem Value="0">all fields</asp:ListItem>
                                    <asp:ListItem Value="1">company name</asp:ListItem>
                                </asp:DropDownList>
                    &nbsp;<asp:Button ID="btn_SearchAddresses" OnClick="btn_SearchAddresses_Click" runat="server"
                        Text="go" ToolTip="click here to begin searching your address book"></asp:Button>
                </td>
                <td align="right">
                    <asp:Button ID="btnBackToDeliveryAddressFromSelectAddress" runat="server" Text="go back"
                        CausesValidation="False" ToolTip="return to address entry" OnClick="btnBackToDeliveryAddressFromSelectAddress_Click">
                    </asp:Button>
                </td>
                <td style="width: 5px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <br />
                    <asp:DataGrid ID="dgAddressBook" runat="server" Width="100%" Font-Size="XX-Small"
                        Font-Names="Verdana" OnPageIndexChanged="dgAddressBook_Page_Change" AllowPaging="True"
                        Visible="False" AutoGenerateColumns="False" GridLines="None" ShowFooter="True"
                        OnItemCommand="dgAddressBook_item_click">
                        <FooterStyle Wrap="False"></FooterStyle>
                        <HeaderStyle Font-Names="Verdana" Wrap="False"></HeaderStyle>
                        <PagerStyle Font-Size="X-Small" Font-Names="Verdana" HorizontalAlign="Center" ForeColor="Blue"
                            BackColor="Silver" Wrap="False" Mode="NumericPages" />
                        <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                        <ItemStyle BackColor="WhiteSmoke"></ItemStyle>
                        <Columns>
                            <asp:TemplateColumn>
                                <ItemStyle Wrap="False" HorizontalAlign="Left"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btnSelectAddress" CommandName="select" Text="select" runat="server"
                                        ToolTip="click here to select this address" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn Visible="False" DataField="DestKey">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Company" HeaderText="Company">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Addr1" HeaderText="Addr 1">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Town" HeaderText="City">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CountryName" HeaderText="Country">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="AttnOf" HeaderText="Attn Of">
                                <HeaderStyle Wrap="False" ForeColor="Navy"></HeaderStyle>
                                <ItemStyle Wrap="False" ForeColor="Navy"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
                <td style="width: 5px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="2">
                    <asp:Label ID="lblAddressMessage" runat="server" ForeColor="Red" Font-Names="Verdana"
                        Font-Size="X-Small"></asp:Label>
                </td>
                <td style="width: 5px">
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlConfirmMultipleAddressBooking" runat="server" Visible="False" Width="100%">
        <table id="Table2" width="100%" runat="server" cellpadding="2" cellspacing="1" style="font-size: x-small;
            font-family: Verdana;">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 10%">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 2%">
                </td>
                <td style="width: 55px">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 5%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <asp:Label ID="Label104" runat="server" Font-Size="Small" ForeColor="Navy">Order: Final Check</asp:Label><br />
                    <br />
                    <asp:Label ID="Label105" runat="server" Font-Size="XX-Small" ForeColor="Navy">Please check the information below before submitting your order for processing. Once we receive your booking we will send you an email confirming our acceptance of your instructions.</asp:Label><br />
                </td>
                <td style="width: 5px">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td valign="top" align="right">
                    <asp:Label ID="Label106" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">From</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblConsignor2" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy"></asp:Label>
                </td>
                <td>
                </td>
                <td style="width: 55px" valign="top">
                    <asp:Label ID="Label107" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">To</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="Label108" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">SEE BELOW</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td align="right" valign="top">
                </td>
                <td align="right" colspan="4">
                    <asp:Button ID="btnConfirmMultipleAddressOrder" runat="server" Text="confirm order"
                        OnClick="btnConfirmMultipleAddressOrder_Click" />
                </td>
                <td>
                </td>
            </tr>
        </table>
        <table id="tabMultipleAddressSummary" runat="server" width="100%" style="font-family: Verdana;
            font-size: xx-small">
        </table>
        <table id="tabMultipleAddressSummaryFooter" runat="server" width="100%" style="font-family: Verdana;
            font-size: xx-small">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 10%">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 2%">
                </td>
                <td style="width: 10%">
                </td>
                <td style="width: 33%">
                </td>
                <td style="width: 5%">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td colspan="5">
                </td>
                <td align="right">
                    <asp:Button ID="btnConfirmMultipleAddressOrder2" runat="server" Text="confirm order"
                        OnClick="btnConfirmMultipleAddressOrder_Click" />
                </td>
                <td>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlConfirmBooking" runat="server" Visible="False" Width="100%">
        <table style="width: 100%; font-size: x-small; font-family: Verdana" cellpadding="2"
            cellspacing="1">
            <tr>
                <td style="width: 5%" />
                <td style="width: 10%" />
                <td style="width: 33%" />
                <td style="width: 2%" />
                <td style="width: 10%" />
                <td style="width: 33%" />
                <td style="width: 5%" />
            </tr>
            <tr runat="server" id="trFinalCheckDefault" visible="true">
                <td />
                <td colspan="5">
                    <asp:Label ID="Label43" runat="server" Font-Size="Small" ForeColor="Navy">Order: Final Check</asp:Label><br />
                    <br />
                    <asp:Label ID="Label44" runat="server" Font-Size="XX-Small" ForeColor="Navy">Please check the information below before submitting your order for processing. Once we receive your order we will send you an email confirming our acceptance of your instructions.</asp:Label><br />
                </td>
                <td />
            </tr>
            <tr runat="server" id="trFinalCheckOrderAuthorisation" visible="true">
                <td />
                <td colspan="5">
                    <asp:Label ID="Label43x" runat="server" Font-Size="Small" ForeColor="Navy">Order: Final Check&nbsp;</asp:Label><asp:Label
                        ID="lblAuthorisationAdvisory01" runat="server" Font-Size="X-Small" ForeColor="Red"
                        Text="&nbsp;&nbsp;&nbsp; (THIS ORDER WILL BE SENT FOR AUTHORISATION)" />
                    <br />
                    <br />
                    <asp:Label ID="lblAuthorisationAdvisory02" runat="server" Font-Size="XX-Small" ForeColor="Navy">Please check the information below before submitting your order for authorisation. You can add an optional message to the authoriser in the box below.</asp:Label><br />
                    <br />
                    <asp:TextBox ID="tbAuthoriserMessageSingleAddressOrder" MaxLength="500" BackColor="lightYellow"
                        Font-Names="Verdana" Font-Size="xX-Small" Width="100%" TextMode="multiLine" runat="server"></asp:TextBox>
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td valign="top" align="right">
                    <asp:Label ID="Label45" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">From</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td valign="top">
                    <asp:Label ID="lblConsignor" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" />
                </td>
                <td />
                <td valign="top" align="right">
                    <asp:Label ID="Label46ytr" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">To</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td valign="top">
                    <asp:Label ID="lblConsignee" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td />
            </tr>
            <tr id="trStandardConfirmationCustRef1" runat="server" visible="true">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendConfirmationCustomerRef1" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy">Reference 1</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblCustRef1" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendConfirmationCustomerRef2" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy">Reference 2</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblCustRef2" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trStandardConfirmationCustRef2" runat="server" visible="true">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendConfirmationCustomerRef3" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy">Reference 3</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblCustRef3" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendConfirmationCustomerRef4" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy">Reference 4</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblCustRef4" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy"></asp:Label>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration1Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label2" runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration1ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration2Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label34" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Booking Ref</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration2ConfirmationBookingRef" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label37" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref A</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefA" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration2Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label38" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref B</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefB" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label39" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref C</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration2ConfirmationAdditionalCustomerRefC" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration3Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label87" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration3ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy"></asp:Label>
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label88" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref A</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefA" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration3Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label89" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref B</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefB" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label90" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref C</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration3ConfirmationAdditionalCustomerRefC" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration4Confirmation1" runat="server" visible="False">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label89zxc" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" Text="Service level: " />
                    <asp:DropDownList ID="ddlPerCustomerConfiguration4Confirmation1ServiceLevel" runat="server"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration4Confirmation1ServiceLevel_Changed"
                        Font-Size="XX-Small" />
                    <asp:CompareValidator ID="cvPerCustomerConfiguration4Confirmation1ServiceLevel" runat="server"
                        ValueToCompare="0" Operator="NotEqual" Font-Names="Verdana" ControlToValidate="ddlPerCustomerConfiguration4Confirmation1ServiceLevel"
                        Font-Size="XX-Small">#</asp:CompareValidator>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration4Confirmation1NotUsed1" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblPerCustomerConfiguration4Confirmation1NotUsed2" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy"></asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration4Confirmation1NotUsed3" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration4Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendPerCustomerConfiguration4Confirmation2TotalValue" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy">Total value for this booking (€):</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration4Confirmation2BasketValue" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendPerCustomerConfiguration4Confirmation2EstimatedFreightCharges"
                        runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy">Estimated freight charges on </asp:Label><asp:Label
                            ID="lblPerCustomerConfiguration4Confirmation2Weight" runat="server" Font-Size="XX-Small"
                            Font-Names="Verdana" ForeColor="Navy"></asp:Label><asp:Label ID="lblLegendPerCustomerConfiguration4Confirmation2Kilos"
                                runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy"> kilos (£):</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration4Confirmation2BasketShippingCost" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" /><asp:HiddenField ID="hidCostCalculationTrace"
                            runat="server" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration4Confirmation3" runat="server" visible="false">
                <td />
                <td colspan="5">
                    <asp:Label ID="lblLegendPerCustomerConfiguration4Confirmation2Return" runat="server"
                        Font-Size="XX-Small" Font-Bold="true" ForeColor="Red" Text="We are unable to accept the return of ordered goods" />
                    <br />
                    <asp:Label ID="lblLegendPerCustomerConfiguration4Confirmation2LocalTaxes" runat="server"
                        Font-Size="XX-Small" Font-Bold="true" ForeColor="Red" Text="Please note: Your order may be subject to additional local duties and taxes not shown here" />
                    <br />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration5Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label96" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration5ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration6Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label49cimax" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Department</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration6ConfirmationDepartment" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label81cimax" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration6ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration6Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label49cimaxreference" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Reference</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration6ConfirmationReference" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                </td>
                <td>
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration7Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label35" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Service Level</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration7ConfirmationServiceLevel" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration18Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label35vsal18" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Service Level</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration18ConfirmationServiceLevel" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration8Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label969" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Booking Ref</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration8ConfirmationBookingRef" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label9698" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">PCID</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration8ConfirmationPCID" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration8Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label9697b" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Rating</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration8ConfirmationRating" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label9697a" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">RO</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration8ConfirmationRO" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration8Confirmation3" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label9696" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">MDS Order Ref</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration8ConfirmationMDSOrderRef" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration12Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Code</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration12ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref A</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefA" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration12Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label140" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref B</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefB" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label141" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Additional Customer Ref C</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration12ConfirmationAdditionalCustomerRefC" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration17Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label35c10qsd17" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Service Level</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration17ConfirmationServiceLevel" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration18Confirmation19" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label219" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration19ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration20Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label220" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration20ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration21Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label220vvf" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre:</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration21ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td>
                    <asp:Label ID="Label51220vvf" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cust Ref:</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration21ConfirmationCustRef" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration22Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138rio" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration22ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139rio" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Requested By</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration22ConfirmationRequestedBy" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration23Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Cost Centre</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration23ConfirmationCostCentre" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Category</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration23ConfirmationCategory" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration23Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label140arth" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">PO Number</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration23ConfirmationPONumber" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td />
                <td />
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration24Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138pv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Reference</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration24ConfirmationReference" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139pv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Requested By</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration24ConfirmationRequestedBy" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration24Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138xpv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Recipient Name</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration24ConfirmationRecipientName" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139xpv" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Products</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration24ConfirmationProducts" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration25Confirmation1" runat="server" visible="False">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label89cab" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy" Text="Service level: " />
                    <asp:DropDownList ID="ddlPerCustomerConfiguration25Confirmation1ServiceLevel" runat="server"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlPerCustomerConfiguration25Confirmation1ServiceLevel_Changed"
                        Font-Size="XX-Small" />
                    <asp:CompareValidator ID="cvPerCustomerConfiguration25Confirmation1ServiceLevel"
                        runat="server" ValueToCompare="0" Operator="NotEqual" Font-Names="Verdana" ControlToValidate="ddlPerCustomerConfiguration25Confirmation1ServiceLevel"
                        Font-Size="XX-Small" Font-Bold="true">###</asp:CompareValidator>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration25Confirmation1NotUsed1" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblPerCustomerConfiguration25Confirmation1NotUsed2" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy"></asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration25Confirmation1NotUsed3" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration25Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendPerCustomerConfiguration25Confirmation2ConsignmentWeight"
                        runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy">Approx weight of this consignment (Kg):</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration25Confirmation2ConsignmentWeight" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendPerCustomerConfiguration25Confirmation2EstimatedFreightCharges"
                        runat="server" Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy">Estimated shipping cost: £</asp:Label><asp:Label
                            ID="lblLegendPerCustomerConfiguration25Confirmation2EstimagedShippingCost" runat="server"
                            Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy"> (£):</asp:Label>
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration25Confirmation2BasketShippingCost" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration29Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label138im" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Budget Code</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration29ConfirmationBudgetCode" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label139im" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Department</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration29ConfirmationDepartment" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration29Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="Label140im" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Event</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration29ConfirmationEvent" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblIMConfBDName" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">BD Name</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration29ConfirmationBDName" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration30Confirmation1" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="LabelJupiter3" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Print Service Level</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration30PrintServiceLevel" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="LabelJupiter4" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Total Print Cost</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfiguration30ConfirmationTotalPrintCost" runat="server"
                        Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr id="trPerCustomerConfiguration30Confirmation2" runat="server" visible="false">
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="LabelJupiter5" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Navy">Budget Code</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td>
                    <asp:Label ID="lblPerCustomerConfirmation30BudgetCode" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" />
                </td>
                <td />
                <td style="white-space: nowrap" align="right">
                </td>
                <td>
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td />
                <td style="white-space: nowrap" align="right">
                    <asp:Label ID="lblLegendSpecialInstructionsConfirmation" runat="server" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Navy" Text="Special Instructions"></asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td colspan="4">
                    <asp:Label ID="lblSpecialInstructions" runat="server" Font-Size="XX-Small" Font-Names="Verdana"
                        ForeColor="Red" />
                </td>
                <td />
            </tr>
            <tr id="trConfirmationPackingNoteText" runat="server" visible="true">
                <td />
                <td align="right" style="white-space: nowrap">
                    <asp:Label ID="Label50ywy" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy">Packing Note Text</asp:Label>&nbsp;&nbsp;&nbsp;
                </td>
                <td colspan="4">
                    <asp:Label ID="lblShippingInfo" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <asp:Label ID="lblLegendStockItems" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy">Stock Items</asp:Label><br />
                    <asp:DataGrid ID="gvConfirmationBasket" runat="server" AutoGenerateColumns="False"
                        CellSpacing="-1" Font-Names="Verdana" Font-Size="XX-Small" GridLines="None" Visible="False"
                        Width="100%">
                        <HeaderStyle ForeColor="Navy" />
                        <AlternatingItemStyle ForeColor="#0000C0" />
                        <ItemStyle ForeColor="Navy" />
                        <Columns>
                            <asp:BoundColumn DataField="ProductCode" HeaderText="Code">
                                <HeaderStyle Font-Underline="True" ForeColor="Navy" Wrap="False" />
                                <ItemStyle ForeColor="Navy" Wrap="False" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ProductDate" HeaderText="Date">
                                <HeaderStyle Font-Underline="True" ForeColor="Navy" Wrap="False" />
                                <ItemStyle ForeColor="Navy" Wrap="False" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Description" HeaderText="Description">
                                <HeaderStyle Font-Underline="True" ForeColor="Navy" />
                                <ItemStyle ForeColor="Navy" Wrap="true" />
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="Qty" Visible="True">
                                <HeaderStyle Font-Underline="True" HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:Label ID="Label48ccprh" runat="server" ForeColor="Navy"><%# Format((DataBinder.Eval(Container, "DataItem.QtyToPick")),"##,##0") %></asp:Label></ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    &nbsp;&nbsp;</ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="Price" Visible="True">
                                <HeaderStyle Font-Underline="True" HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:Label ID="Label48ccprh" runat="server" ForeColor="Navy" Visible='<%# sSetIsCABVisibility %>'><%# Format((DataBinder.Eval(Container, "DataItem.UnitValue")),"##,##0.00") %></asp:Label></ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="Total" Visible="True">
                                <HeaderStyle Font-Underline="True" HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                                <ItemTemplate>
                                    <asp:Label ID="Label48" runat="server" ForeColor="Navy" Visible='<%# sSetIsCABVisibility %>'><%# Format((DataBinder.Eval(Container, "DataItem.QtyToPick")) * (DataBinder.Eval(Container, "DataItem.UnitValue")),"##,##0.00") %></asp:Label></ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
                <td />
            </tr>
            <tr id="trMaterialsTotal" runat="server" visible='<%# sSetIsCABVisibility %>'>
                <td />
                <td colspan="5" align="right">
                    <asp:Label ID="lblMaterialsSummary" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Navy" />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5" align="right">
                    <asp:LinkButton ID="lnkbtnSaveAddressInPersonalAddressBook" runat="server" OnClick="lnkbtnSaveAddressInPersonalAddressBook_Click">save address in personal address book</asp:LinkButton>&nbsp;&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnSaveAddressInSharedAddressBook" runat="server" OnClick="lnkbtnSaveAddressInSharedAddressBook_Click">save address in shared address book</asp:LinkButton>&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnFinalCheckGoBack" runat="server" CausesValidation="false" OnClick="btn_CheckOut_click"
                        Text="go back" ToolTip="go back and modify your order" />
                    &nbsp;&nbsp;
                    <asp:Button ID="btnViewInvoice" runat="server" Text="view invoice" ToolTip="view your invoice" />
                    &nbsp;&nbsp;
                    <asp:Button ID="btnSingleAddressConfirmOrder" runat="server" OnClick="btnSingleAddressConfirmOrder_click"
                        Text="confirm order" ToolTip="click here to submit your order" />
                </td>
                <td />
            </tr>
            <tr>
                <td />
                <td colspan="5">
                    <hr />
                </td>
                <td />
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlBookingConfirmation" runat="server" Visible="False" Width="100%">
        <asp:Table ID="Table1" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana"
            CellPadding="2" CellSpacing="1">
            <asp:TableRow runat="server">
                <asp:TableCell Width="5%" runat="server"></asp:TableCell><asp:TableCell Width="10%"
                    runat="server"></asp:TableCell><asp:TableCell Width="33%" runat="server"></asp:TableCell><asp:TableCell
                        Width="2%" runat="server"></asp:TableCell><asp:TableCell Width="10%" runat="server"></asp:TableCell><asp:TableCell
                            Width="33%" runat="server"></asp:TableCell><asp:TableCell Width="5%" runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server"></asp:TableCell><asp:TableCell ColumnSpan="5" runat="server">
                    <asp:Label ID="Label6" runat="server" Font-Size="Small" ForeColor="Navy">Order: Complete</asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="Label7" runat="server" Font-Size="XX-Small" ForeColor="Navy">Thank you for your order, which is now being processed. You should shortly receive a confirmation by e-mail.</asp:Label>
                </asp:TableCell><asp:TableCell runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow runat="server">
                <asp:TableCell runat="server"></asp:TableCell><asp:TableCell ColumnSpan="2" runat="server">
                    <asp:Label ID="Label12" runat="server" ForeColor="Navy" Font-Names="Verdana" Font-Size="X-Small">Consignment No</asp:Label>&nbsp;&nbsp;
                </asp:TableCell><asp:TableCell ColumnSpan="3" runat="server">
                    <asp:Label ID="lblConsignmentNo" runat="server" ForeColor="Red" Font-Names="Verdana"
                        Font-Size="X-Small"></asp:Label>
                </asp:TableCell><asp:TableCell runat="server"></asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlBookingQueuedConfirmation" runat="server" Visible="False" Width="100%">
        <asp:Table ID="Table4" runat="server" Width="100%" Font-Size="X-Small" Font-Names="Verdana"
            CellPadding="2" CellSpacing="1">
            <asp:TableRow ID="TableRow1" runat="server">
                <asp:TableCell ID="TableCell36" Width="5%" runat="server"></asp:TableCell><asp:TableCell
                    ID="TableCell37" Width="10%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell38"
                        Width="33%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell40" Width="2%"
                            runat="server"></asp:TableCell><asp:TableCell ID="TableCell41" Width="10%" runat="server"></asp:TableCell><asp:TableCell
                                ID="TableCell69" Width="33%" runat="server"></asp:TableCell><asp:TableCell ID="TableCell70"
                                    Width="5%" runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow ID="TableRow6" runat="server">
                <asp:TableCell ID="TableCell71" runat="server"></asp:TableCell><asp:TableCell ID="TableCell72"
                    ColumnSpan="5" runat="server">
                    <asp:Label ID="Label42" runat="server" Font-Size="Small" ForeColor="Navy">Order: Sent for authorisation</asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="lblAuthorisationAdvisory03" runat="server" Font-Size="XX-Small" ForeColor="Navy">Thank you for your order, which has been sent for authorisation.</asp:Label>
                </asp:TableCell><asp:TableCell ID="TableCell73" runat="server"></asp:TableCell></asp:TableRow>
            <asp:TableRow ID="TableRow12" runat="server">
                <asp:TableCell ID="TableCell74" runat="server"></asp:TableCell><asp:TableCell ID="TableCell75"
                    ColumnSpan="2" runat="server">
                </asp:TableCell><asp:TableCell ID="TableCell76" ColumnSpan="3" runat="server">
                </asp:TableCell><asp:TableCell ID="TableCell77" runat="server"></asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlClassicProductList" runat="server" Visible="False" Width="100%">
        <asp:Table ID="Table5" runat="Server" Width="90%">
            <asp:TableRow>
                <asp:TableCell Wrap="False" ColumnSpan="2">
                    <asp:Label runat="server" ID="lblClassicCategoryHeader" Font-Names="Verdana" ForeColor="Blue"
                        Font-Bold="True" Font-Size="Small"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell VerticalAlign="Bottom" Width="5%">
                    <asp:Image ID="Image1" runat="server" ImageUrl="./images/icon_back.gif"></asp:Image>
                </asp:TableCell><asp:TableCell Wrap="False" VerticalAlign="Top">
                    <asp:LinkButton ID="LinkButton7" runat="server" ForeColor="Blue" Font-Size="X-Small"
                        Font-Names="Verdana" OnClick="btn_ClassicReSelectCategory_click">re-select category or search</asp:LinkButton>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
        </asp:Table>
        <asp:Label ID="lblClassicProductMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="#00C000"></asp:Label><asp:DataGrid ID="dgrdProducts" runat="server" Font-Size="XX-Small"
                Width="100%" Font-Names="Verdana" OnItemCommand="ClassicProductGrid_item_click"
                AutoGenerateColumns="False" OnSortCommand="SortProductColumns" AllowSorting="True"
                ShowFooter="True" GridLines="None" CellSpacing="1" CellPadding="1" OnItemDataBound="dgrdProducts_ItemDataBound">
                <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" Wrap="False" BorderColor="Gray">
                </HeaderStyle>
                <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                <ItemStyle Font-Names="Verdana" BackColor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticProductKey"></asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle ForeColor="Blue" Width="5%"></HeaderStyle>
                        <ItemStyle VerticalAlign="Top"></ItemStyle>
                        <HeaderTemplate>
                            Info</HeaderTemplate>
                        <ItemTemplate>
                            <asp:ImageButton ID="imgBtnShowProdDetails" runat="server" CommandName="info" ImageUrl="./images/icon_info.gif"
                                ToolTip="product info" />
                            <asp:HiddenField ID="hidClassicCalendarManaged1" Value='<%# DataBinder.Eval(Container.DataItem,"CalendarManaged") %>'
                                runat="server" />
                            <asp:HiddenField ID="hidClassicCustomLetter1" Value='<%# DataBinder.Eval(Container.DataItem,"CustomLetter") %>'
                                runat="server" />
                            <asp:HiddenField ID="hidOnDemand1" Value='<%# DataBinder.Eval(Container.DataItem,"OnDemand") %>'
                                runat="server" />
                            <asp:HiddenField ID="hidOnDemandPriceList1" Value='<%# DataBinder.Eval(Container.DataItem,"OnDemandPriceList") %>'
                                runat="server" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Code">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn Visible="False" DataField="ProductDate" SortExpression="ProductDate"
                        HeaderText="Date">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription"
                        HeaderText="Description">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn Visible="False" DataField="ProductDepartmentId" SortExpression="ProductDepartmentId"
                        HeaderText="Category">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="LanguageId" SortExpression="LanguageId" HeaderText="Language">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="UnitValue" SortExpression="UnitValue" HeaderText="Value"
                        DataFormatString="{0:#,##0.00}">
                        <HeaderStyle Wrap="False" HorizontalAlign="Right" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn Visible="False" DataField="ItemsPerBox" SortExpression="ItemsPerBox"
                        HeaderText="Box Qty">
                        <HeaderStyle Wrap="False"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn Visible="False" DataField="UnitWeightGrams">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Quantity" SortExpression="Quantity" HeaderText="Quantity"
                        DataFormatString="{0:#,##0}">
                        <HeaderStyle Wrap="False" HorizontalAlign="Right" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Top"></ItemStyle>
                        <HeaderTemplate>
                            <asp:Button ID="btnClassicViewAddToBasket" Text="add to basket" runat="server" OnClick="btn_AddToBasket_click"
                                Font-Names="Verdana" Font-Size="XX-Small"></asp:Button>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:CheckBox ID="CheckBox1" runat="server"></asp:CheckBox>
                        </ItemTemplate>
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                        <FooterTemplate>
                            <asp:Button ID="Button5" Text="add to basket" runat="server" OnClick="btn_AddToBasket_click"
                                Font-Names="Verdana" Font-Size="XX-Small"></asp:Button>
                        </FooterTemplate>
                    </asp:TemplateColumn>
                </Columns>
            </asp:DataGrid>
        <asp:Table ID="Table6" runat="Server" Font-Size="X-Small" Width="100%" Font-Names="Verdana">
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell HorizontalAlign="Right">
                    <asp:LinkButton ID="LinkButton8" runat="server" ForeColor="Blue" Font-Size="XX-Small"
                        Font-Names="Verdana" OnClick="btn_RefreshProdList_click" Visible="False">refresh</asp:LinkButton>
                </asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlClassicBasket" runat="server" Visible="False" Width="100%">
        <asp:Label ID="Label100" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="#0000C0" Font-Bold="True">&nbsp;Your Basket</asp:Label><br />
        <br />
        <asp:Label ID="lblClassicBasketMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="#00C000"></asp:Label><div align="left">
                <asp:DataGrid ID="dgrdBasket" runat="server" Font-Size="XX-Small" Width="100%" Font-Names="Verdana"
                    AutoGenerateColumns="False" OnSortCommand="SortBasketColumns" AllowSorting="True"
                    ShowFooter="True" GridLines="None" Visible="False" CellSpacing="1" CellPadding="1"
                    OnItemDataBound="dgrdBasket_ItemDataBound">
                    <HeaderStyle Font-Size="XX-Small"></HeaderStyle>
                    <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                    <ItemStyle BackColor="LightGray"></ItemStyle>
                    <Columns>
                        <asp:BoundColumn Visible="False" DataField="ProductKey" ReadOnly="True" />
                        <asp:TemplateColumn HeaderText="Remove">
                            <HeaderStyle HorizontalAlign="Center" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" HorizontalAlign="Center" VerticalAlign="Top"></ItemStyle>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkRemove" runat="server"></asp:CheckBox><asp:HiddenField ID="hidClassicCalendarManaged2"
                                    Value='<%# DataBinder.Eval(Container.DataItem,"CalendarManaged") %>' runat="server" />
                            </ItemTemplate>
                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                            <FooterTemplate>
                                <asp:Button ID="btnRemove" runat="server" CausesValidation="False" OnClick="btn_ClassicRemoveBasketItems_click"
                                    Text="remove" Font-Names="Verdana" Font-Size="XX-Small" ToolTip="remove checked items above">
                                </asp:Button>
                            </FooterTemplate>
                        </asp:TemplateColumn>
                        <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Code">
                            <HeaderStyle ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="ProductDate" SortExpression="ProductDate"
                            HeaderText="Date">
                            <HeaderStyle ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Description" SortExpression="Description" HeaderText="Description">
                            <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="LanguageID" SortExpression="Language" HeaderText="Language">
                            <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                            <FooterStyle Wrap="False"></FooterStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="UnitValue" SortExpression="UnitValue" HeaderText="Unit Value"
                            DataFormatString="{0:#,##0.00}">
                            <HeaderStyle Wrap="False" HorizontalAlign="Right" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="BoxQty" SortExpression="BoxQty" HeaderText="Box Qty">
                            <HeaderStyle Wrap="False" HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle Wrap="False" HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="QtyAvailable" SortExpression="QtyAvailable" HeaderText="Qty Available"
                            DataFormatString="{0:#,##0}">
                            <HeaderStyle Wrap="False" HorizontalAlign="Right" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" HorizontalAlign="Right" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="UnitWeightGrams">
                            <HeaderStyle ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" VerticalAlign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn>
                            <HeaderStyle HorizontalAlign="Right" ForeColor="Blue"></HeaderStyle>
                            <ItemStyle Wrap="False" HorizontalAlign="Right"></ItemStyle>
                            <HeaderTemplate>
                                Qty to Pick</HeaderTemplate>
                            <ItemTemplate>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" Font-Bold="True" Font-Names="Verdana"
                                    ForeColor="Red" runat="Server" Text=">>>" ControlToValidate="txtPickQuantity"></asp:RequiredFieldValidator><asp:TextBox
                                        ID="txtPickQuantity" Font-Size="XX-Small" Width="40px" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.QtyToPick") %>'></asp:TextBox></ItemTemplate>
                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
            </div>
        <asp:Table ID="Table7" runat="Server" Font-Size="X-Small" Width="100%" Font-Names="Verdana">
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell HorizontalAlign="Right">
                    <asp:LinkButton runat="server" ForeColor="Blue" ID="lnkbtnClassicProceedToCheckout"
                        Font-Size="XX-Small" OnClick="btn_CheckOut_click">proceed to checkout</asp:LinkButton>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell HorizontalAlign="Right">
                    <asp:LinkButton ID="LinkButton9" runat="server" ForeColor="Blue" CausesValidation="False"
                        Font-Size="XX-Small" OnClick="btn_ReturnToProducts_click">go back to product list</asp:LinkButton>
                </asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlClassicProductDetail" runat="server" Visible="False">
        <asp:Label ID="Label116" runat="server" Font-Names="Verdana" Font-Size="X-Small"
            ForeColor="#0000C0" Font-Bold="True">Product
            details</asp:Label><br />
        <br />
        <asp:Table ID="Table13" runat="server" Font-Size="XX-Small" Width="750px" Font-Names="Verdana">
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell ColumnSpan="4" HorizontalAlign="Right"></asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell Width="10px"></asp:TableCell><asp:TableCell VerticalAlign="Top" RowSpan="12"
                    Width="100px">
                    <asp:HyperLink ID="hlnk_DetailThumb" runat="server" Target="_blank" ToolTip="click here to see larger image"></asp:HyperLink>
                </asp:TableCell><asp:TableCell Width="10px"></asp:TableCell><asp:TableCell BackColor="Lavender"
                    Width="150px" HorizontalAlign="Left">
                    <asp:Label ID="Label117" runat="server" ForeColor="#0000C0">Product Code</asp:Label>
                </asp:TableCell><asp:TableCell Width="479px" BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblProductCode"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label118" runat="server" ForeColor="#0000C0">Product Date</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblProductDate"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label119" runat="server" ForeColor="#0000C0">Description</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblDescription"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label120" runat="server" ForeColor="#0000C0">Language</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblLanguage"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label121" runat="server" ForeColor="#0000C0">Cost Centre</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblDepartment"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label122" runat="server" ForeColor="#0000C0">Category</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblCategory"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label123" runat="server" ForeColor="#0000C0">Sub Category</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblSubCategory"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label124" runat="server" ForeColor="#0000C0">Items Per Box</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblItemsPerBox"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label125sle" runat="server" ForeColor="#0000C0">Min Stock Level</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblMinStockLevel"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="lblLegendUnitValue" runat="server" ForeColor="#0000C0" Text="Unit Value (£)" />
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblUnitValue"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell></asp:TableCell><asp:TableCell></asp:TableCell><asp:TableCell BackColor="Lavender">
                    <asp:Label ID="Label127" runat="server" ForeColor="#0000C0">Unit Weight (grams)</asp:Label>
                </asp:TableCell><asp:TableCell BackColor="Gainsboro">
                    <asp:Label runat="server" ID="lblUnitWeight"></asp:Label>
                </asp:TableCell></asp:TableRow>
            <asp:TableRow>
                <asp:TableCell ColumnSpan="3">
                    <asp:HyperLink runat="Server" ID="hyplnkPDFDocument" Visible="False" ForeColor="Blue"
                        Target="_blank" Text="view PDF document" />
                </asp:TableCell><asp:TableCell HorizontalAlign="right" ColumnSpan="2">
                    <br />
                    <asp:LinkButton runat="server" ID="btnBackToProductList" OnClick="btn_BackToProductList_click"
                        ForeColor="Blue">back to product list</asp:LinkButton>
                </asp:TableCell><asp:TableCell></asp:TableCell></asp:TableRow>
        </asp:Table>
    </asp:Panel>
    <asp:Panel ID="pnlCalendarManaged" runat="server" Visible="False" Width="100%">
        <strong>&nbsp;Calendar Managed Products</strong><br />
        <table style="width: 100%">
            <tr>
                <td style="width: 10%">
                </td>
                <td style="width: 30%">
                </td>
                <td style="width: 60%" align="right">
                    <asp:Button ID="ButtonCM2" runat="server" Text="back to product list" OnClick="btn_ReturnToProducts_click" />
                    <asp:Button ID="ButtonCM1" runat="server" Text="back to basket" OnClick="btn_ViewCurrentBasket_click" />
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" valign="top">
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 35%" valign="top">
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 42%" valign="top">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td colspan="5">
                    <asp:GridView ID="gvCalendarManagedItems" runat="server" Font-Names="Verdana" Width="100%"
                        Font-Size="XX-Small" CellPadding="2" AutoGenerateColumns="False">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:HyperLink ID="hlnk_ThumbNail2" runat="server" ToolTip="click here to see larger image"
                                        NavigateUrl='<%# "Javascript:SB_ShowImage(""" & DataBinder.Eval(Container.DataItem,"OriginalImage") & """)" %> '
                                        ImageUrl='<%# psVirtualThumbFolder & DataBinder.Eval(Container.DataItem,"ThumbNailImage") & "?" & Now.ToString %>'></asp:HyperLink></ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Label ID="lblCalendarManagedItemProductCode" Text='<%# DataBinder.Eval(Container.DataItem,"ProductCode") %>'
                                        runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Label ID="lblCalendarManagedItemVersionDate" Text='<%# DataBinder.Eval(Container.DataItem,"ProductDate") %>'
                                        runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Label ID="lblCalendarManagedItemDescription" Text='<%# DataBinder.Eval(Container.DataItem,"Description") %>'
                                        runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label></ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" valign="top" align="center">
                    <asp:Label ID="lblCMLegendCalendar" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="XX-Small" />
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 35%" valign="top" align="center">
                    <strong>
                        <asp:Label ID="Label53" runat="server" Font-Names="Verdana" Font-Size="X-Small" Text="Enter event details"></asp:Label></strong>
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 42%" valign="top" align="center">
                    <asp:Label ID="lblCMLegendViewOtherReservations" runat="server" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
            </tr>
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 20%" valign="top">
                    <asp:Calendar ID="calCalendar1" runat="server" OnSelectionChanged="calCalendar1_SelectionChanged"
                        OnDayRender="calCalendar1_DayRender" CellPadding="5" OnVisibleMonthChanged="calCalendar1_VisibleMonthChanged"
                        Font-Names="Verdana" Font-Size="XX-Small" />
                    <table style="width: 100%">
                        <tr>
                            <td style="width: 50%">
                            </td>
                            <td style="width: 50%">
                            </td>
                        </tr>
                        <tr>
                            <td align="right" style="white-space: nowrap">
                                <asp:Label ID="Label46" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Delivery date:"></asp:Label>
                            </td>
                            <td>
                                <asp:Label ID="lblCMDeliveryDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" style="white-space: nowrap">
                                <asp:Label ID="Label50" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Collection date:"></asp:Label>
                            </td>
                            <td>
                                <asp:Label ID="lblCMCollectionDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <asp:LinkButton ID="lnkbtnCMClearDateSelection" runat="server" OnClick="lnkbtnCMClearDateSelection_Click"
                                    Font-Names="Verdana" Font-Size="XX-Small">clear date selection</asp:LinkButton><br />
                                <asp:LinkButton ID="lnkbtnFindAvailableProducts" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnFindAvailableProducts_Click" Visible="False">find available products</asp:LinkButton>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" style="height: 181px">
                                <table style="width: 100%">
                                    <tr>
                                        <td style="width: 25%" align="center">
                                            <asp:Label ID="Label87sdfg" runat="server" Font-Bold="True" Font-Names="Verdana"
                                                Font-Size="XX-Small" Text="Key:"></asp:Label>
                                        </td>
                                        <td style="width: 75%">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 21px" align="center">
                                            &nbsp;<asp:Label ID="Label95" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="white" />
                                        </td>
                                        <td style="height: 25px">
                                            <asp:Label ID="Label88dfgh" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="product(s) can be booked for this date" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="background-color: dimgray" align="center">
                                            <asp:Label ID="Label97" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="grey" />
                                        </td>
                                        <td style="height: 25px">
                                            <asp:Label ID="Label89fgjh" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="date in the past" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 21px; background-color: lightgrey" align="center">
                                            <asp:Label ID="Label102" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="silver" />
                                        </td>
                                        <td style="height: 25px">
                                            <asp:Label ID="Label90ghjk" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="please confirm booking by telephone" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="background-color: red" align="center">
                                            <asp:Label ID="Label103bnm" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="red" />
                                        </td>
                                        <td style="height: 25px">
                                            <asp:Label ID="Label94" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="selected product(s) already booked for this date" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" style="background-color: green">
                                            <asp:Label ID="Label112" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="green" />
                                        </td>
                                        <td style="height: 25px">
                                            <asp:Label ID="Label96cvb" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                                Text="date selected" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 35%" valign="top">
                    <table style="width: 100%">
                        <tr>
                            <td style="width: 30%" align="right">
                                <asp:RequiredFieldValidator ID="rfvCMEventName" ControlToValidate="tbCMEventName"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <asp:Label ID="Label2axa" runat="server" ForeColor="Red" Text="Event name:" Font-Names="Verdana"
                                    Font-Size="XX-Small"></asp:Label>
                            </td>
                            <td style="width: 70%">
                                <asp:TextBox ID="tbCMEventName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactName" ControlToValidate="tbCMContactName"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small"></asp:RequiredFieldValidator><asp:Label ID="Label6axa" runat="server"
                                        ForeColor="Red" Text="Contact name:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMContactName" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactPhone" ControlToValidate="tbCMContactPhone"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <asp:Label ID="Label15axa" runat="server" ForeColor="Red" Text="Contact phone:" Font-Names="Verdana"
                                    Font-Size="XX-Small"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMContactPhone" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactMobile" ControlToValidate="tbCMContactMobile"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <asp:Label ID="Label34axa" runat="server" Text="Contact mobile:" ForeColor="Red"
                                    Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMContactMobile" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr id="trCMContactName2" runat="server" visible="true">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactName2" runat="server" ControlToValidate="tbCMContactName2"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged"
                                    Visible="False" />
                                <asp:Label ID="lblLegendCMContactName2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Contact name 2:" Visible="False" />
                            </td>
                            <td>
                                <asp:LinkButton ID="lnkbtnCMAddSecondContact" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMAddSecondContact_Click">add 2nd contact</asp:LinkButton><asp:TextBox
                                        ID="tbCMContactName2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                        MaxLength="50" Width="100%" Visible="False" />
                            </td>
                        </tr>
                        <tr id="trCMContactPhone2" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactPhone2" runat="server" ControlToValidate="tbCMContactPhone2"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label15axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Contact phone 2:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMContactPhone2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="100%" />
                            </td>
                        </tr>
                        <tr id="trCMContactMobile2" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMContactMobile2" runat="server" ControlToValidate="tbCMContactMobile2"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label34axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Contact mobile 2:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMContactMobile2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="100%" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                &nbsp;
                            </td>
                            <td>
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMPostcode" runat="server" ControlToValidate="tbCMPostcode"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label39axa" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Event postcode:" ForeColor="Red" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="80px" />
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnCMFindEventAddress" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMFindEventAddress_Click">find 
                                    addr</asp:LinkButton>&nbsp;<asp:Label ID="lblCMFindEventAddressFailure" runat="server"
                                        Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="no results found  !!"
                                        Visible="False" />
                                &nbsp;<asp:LinkButton ID="lnkbtnCMAddressOutsideUK" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMAddressOutsideUK_Click">addr 
                                    outside UK</asp:LinkButton>
                            </td>
                        </tr>
                        <tr id="trCMSelectEventAddress" runat="server" visible="false">
                            <td align="right">
                                <asp:Label ID="Label79" runat="server" Text="Select address:" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <br />
                                <br />
                                <asp:LinkButton ID="lnkbtnCMCancelSelectEventAddress" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMCancelSelectEventAddress_Click">cancel</asp:LinkButton>
                            </td>
                            <td>
                                <asp:ListBox ID="lbCMSelectEventAddress" runat="server" AutoPostBack="True" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnSelectedIndexChanged="lbCMSelectEventAddress_SelectedIndexChanged"
                                    Rows="8" Width="100%" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMEventAddress1" runat="server" ControlToValidate="tbCMEventAddress1"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label35axa" runat="server" Text="Event addr 1:" Font-Names="Verdana"
                                    Font-Size="XX-Small" ForeColor="Red" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMEventAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:Label ID="Label37axa" runat="server" Text="Event addr 2:" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMEventAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="100%" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMTown" runat="server" ControlToValidate="tbCMTown"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label38axa" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Town/City:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="100%" />
                            </td>
                        </tr>
                        <tr id="trCMCountry" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCountry" runat="server" ControlToValidate="ddlCMCountry"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label38axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Country:" />
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlCMCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Navy" TabIndex="8" Width="100%">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                            </td>
                            <td>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMItemDeliverBy" ControlToValidate="ddlCMItemDeliverBy"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" InitialValue="- please select -" />
                                <asp:Label ID="Label42axa" runat="server" ForeColor="Red" Text="Deliver by:" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlCMItemDeliverBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                                    <asp:ListItem>- please select -</asp:ListItem>
                                    <asp:ListItem>9.00am</asp:ListItem>
                                    <asp:ListItem>10.30am</asp:ListItem>
                                    <asp:ListItem>12.00 noon</asp:ListItem>
                                    <asp:ListItem Value="Other times pls specify in Special Instructions">Other 
                                        times pls specify in Special Instructions</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMExactDeliveryPoint" runat="server" ControlToValidate="tbCMExactDeliveryPoint"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label43axa" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Exact delivery point:"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMExactDeliveryPoint" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="100" Width="100%" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right" style="height: 22px">
                            </td>
                            <td style="height: 22px">
                                <asp:CheckBox ID="cbCMDifferentCollectionAddress" runat="server" AutoPostBack="True"
                                    Font-Names="Verdana" Font-Size="XX-Small" OnCheckedChanged="cbCMDifferentCollectionAddress_CheckedChanged"
                                    Text="collect from a different address" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMItemCollectBetween" ControlToValidate="ddlCMItemCollectBetween"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" InitialValue="- please select -" />
                                <asp:Label ID="Label44axa" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Collect between:"></asp:Label>
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlCMItemCollectBetween" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small">
                                    <asp:ListItem>- please select -</asp:ListItem>
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
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trCMCollectionPostcode" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCollectionPostcode" runat="server" ControlToValidate="tbCMCollectionPostcode"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label78" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Collect'n postcode:" ForeColor="Red" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMCollectionPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="80px" />
                                &nbsp;
                                <asp:LinkButton ID="lnkbtnCMFindCollectionAddress" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMFindCollectionAddress_Click">find addr</asp:LinkButton>&nbsp;<asp:Label
                                        ID="lblCMFindCollectionAddressFailure" runat="server" Font-Bold="True" Font-Names="Verdana"
                                        Font-Size="XX-Small" ForeColor="Red" Text="no results found  !!" Visible="False"></asp:Label>&nbsp;<asp:LinkButton
                                            ID="lnkbtnCMCollectionAddressOutsideUK" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                            OnClick="lnkbtnCMCollectionAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                            </td>
                        </tr>
                        <tr id="trCMSelectCollectionAddress" runat="server" visible="false">
                            <td align="right" style="height: 59px">
                                <asp:Label ID="Label80" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Select address:"></asp:Label><br />
                                <br />
                                <asp:LinkButton ID="lnkbtnCMCancelSelectCollectionAddress" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" OnClick="lnkbtnCMCancelSelectCollectionAddress_Click">cancel</asp:LinkButton>
                            </td>
                            <td style="height: 59px">
                                <asp:ListBox ID="lbCMSelectCollectionAddress" runat="server" AutoPostBack="True"
                                    Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="lbCMSelectCollectionAddress_SelectedIndexChanged"
                                    Rows="8" Width="100%" />
                            </td>
                        </tr>
                        <tr id="trCMCollectionAddress1" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCollectionAddress1" runat="server" ControlToValidate="tbCMCollectionAddress1"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label54" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Collect'n addr 1:" ForeColor="Red" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMCollectionAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr id="trCMCollectionAddress2" runat="server" visible="false">
                            <td align="right">
                                &nbsp;<asp:Label ID="Label55tyt" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="Collect'n addr 2:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMCollectionAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr id="trCMCollectionTown" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCollectionTown" ControlToValidate="tbCMCollectionTown"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <asp:Label ID="Label77" runat="server" ForeColor="Red" Text="Town/City:" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMCollectionTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="50" />
                            </td>
                        </tr>
                        <tr id="trCMCollectionCountry" runat="server" visible="false">
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCollectionCountry" runat="server" ControlToValidate="ddlCMCollectionCountry"
                                    ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="Label38axa1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    ForeColor="Red" Text="Country:" />
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlCMCollectionCountry" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="100%">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMExactCollectionPoint" ControlToValidate="tbCMExactCollectionPoint"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" />
                                <asp:Label ID="Label45axa" runat="server" Text="Exact collection point:" Font-Names="Verdana"
                                    Font-Size="XX-Small" ForeColor="Red" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMExactCollectionPoint" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="100" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMSpecialInstructions" ControlToValidate="tbCMSpecialInstructions"
                                    runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged" Font-Names="Verdana"
                                    Font-Size="XX-Small" Visible="False" />
                                <asp:Label ID="lblLegendCMSpecialInstructions" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" Text="Special instructions:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMSpecialInstructions" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Width="100%" MaxLength="180" Rows="2" TextMode="MultiLine" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right">
                                <asp:RequiredFieldValidator ID="rfvCMCustomerReference" runat="server" ControlToValidate="tbCMCustomerReference"
                                    Enabled="False" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                                <asp:Label ID="lblLegendCMCustomerReference" runat="server" Font-Names="Verdana"
                                    Font-Size="XX-Small" Text="Customer reference:" />
                            </td>
                            <td>
                                <asp:TextBox ID="tbCMCustomerReference" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    MaxLength="50" Width="100%" />
                            </td>
                        </tr>
                        <tr>
                            <td align="right" colspan="2">
                                <asp:Button ID="btnCMBookEvent" runat="server" OnClick="btnCMBookEvent_Click" Text="book items for event" />
                                <asp:CheckBox ID="cbCMMultipleBookings" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                                    Text="multiple bookings" ToolTip="Select this check box to make more than one booking for the same item(s)" />
                            </td>
                        </tr>
                    </table>
                </td>
                <td style="width: 1%">
                </td>
                <td style="width: 42%" valign="top">
                    <br />
                    <asp:GridView ID="gvOtherCalendarManagedReservations" runat="server" CellPadding="2"
                        Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                        <EmptyDataTemplate>
                            no events found</EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
        </table>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlCustomLetter" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
    </asp:Panel>
    <asp:Panel ID="pnlFindAvailableProducts" runat="server" Font-Names="Verdana" Font-Size="X-Small"
        Width="100%" Visible="false">
        <table style="width: 100%">
            <tr>
                <td style="width: 50%">
                    <asp:Label ID="Label48ccldfgwe" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text="Matching Available Products - " />
                    <asp:Label ID="lblCMAvailableProductsFromDate" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" />
                    <asp:Label ID="Label48ccldfgwe0" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" Text=" to " />
                    <asp:Label ID="lblCMAvailableProductsToDate" runat="server" Font-Bold="True" Font-Names="Verdana"
                        Font-Size="X-Small" />
                </td>
                <td style="width: 50%" align="right">
                    <asp:Button ID="btnBackFromFindAvailableProductsToEvent" runat="server" Text="back to event"
                        OnClick="btnBackFromFindAvailableProductsToEvent_Click" />
                </td>
            </tr>
        </table>
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                </td>
                <td style="width: 95%">
                    <asp:GridView ID="gvCMAvailableProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoGenerateColumns="False" CellPadding="2" Width="95%">
                        <Columns>
                            <asp:TemplateField HeaderText="add to basket" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbCMAddAvailableProductToBasket" runat="server" Font-Names="Verdana"
                                        Font-Size="XX-Small" />
                                    <asp:HiddenField ID="hidAvailableProductKey" runat="server" Value='<%# DataBinder.Eval(Container.DataItem,"LogisticProductKey") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True"
                                SortExpression="ProductCode" />
                            <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True"
                                SortExpression="ProductDescription" />
                            <asp:BoundField DataField="LanguageId" HeaderText="Type" ReadOnly="True" SortExpression="LanguageId" />
                        </Columns>
                        <EmptyDataTemplate>
                            no available products found for the selected dates(s)</EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnCMAddAvailableProductsToBasket" runat="server" Text="add products to basket"
                        OnClick="btnCMAddAvailableProductsToBasket_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="cbCMRemoveNonBookableProductsFromBasket" runat="server" Checked="True"
                        Font-Names="Verdana" Font-Size="XX-Small" Text="remove non-bookable products from basket" />
                </td>
            </tr>
        </table>
    </asp:Panel>
    &nbsp;&nbsp;<asp:Label ID="lblError" runat="server" ForeColor="Red" Font-Names="Verdana"
        Font-Size="X-Small"></asp:Label></form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
