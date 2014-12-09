<%@ Control Language="VB" %>
<%@ Import Namespace="System.Net" %>
<%@ Register TagPrefix="VASP" Assembly="VASPTBv4NET" Namespace="VASPTBv4NET" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.XML" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.ServiceModel.Syndication" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<script runat="server">

    Const CONFIG_XML_FILE As String = "news_config.xml"
    Const COLOUR_LIGHT_GREY As String = "#cccccc"
    Const COLOUR_GREY As String = "#adadad"
    Const COLOUR_WHITE As String = "#ffffff"
    Const COLOUR_NAVY_BLUE As String = "#000A61"  ' SEAC

    Const USER_PERMISSION_A_NOBODY As Integer = 0

    Const USER_PERMISSION_ACCOUNT_HANDLER As Integer = 1
    Const USER_PERMISSION_SITE_ADMINISTRATOR As Integer = 2
    Const USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR As Integer = 4
    Const USER_PERMISSION_SITE_EDITOR As Integer = 8

    Const USER_PERMISSION_DEPUTY_SITE_EDITOR As Integer = 16
    Const USER_PERMISSION_32_NOT_USED As Integer = 32
    Const USER_PERMISSION_64_NOT_USED As Integer = 64
    Const USER_PERMISSION_128_NOT_USED As Integer = 128
    
    Const USER_PERMISSION_256_NOT_USED As Integer = 256
    Const USER_PERMISSION_512_NOT_USED As Integer = 512
    Const USER_PERMISSION_VIEW_STOCK As Integer = 1024
    Const USER_PERMISSION_CREATE_STOCK_BOOKING As Integer = 2048

    Const USER_PERMISSION_PRINT_ON_DEMAND_TAB As Integer = 4096
    Const USER_PERMISSION_ADVANCED_PERMISSIONS_TAB As Integer = 8192
    Const USER_PERMISSION_FILE_UPLOAD_TAB As Integer = 16384
    Const USER_PERMISSION_WU_INTERNAL_USER As Integer = &H8000

    Const USER_PERMISSION_WU_CREATE_DELETE_ACCOUNTS As Integer = &H10000
    Const USER_PERMISSION_WU_RESET_PASSWORDS As Integer = &H20000
    Const USER_PERMISSION_WU_ACCESS_REPORTS As Integer = &H40000
    Const USER_PERMISSION_WU_VIEW_STOCK As Integer = &H80000

    Const USER_PERMISSION_WU_ORDER_STOCK As Integer = &H100000
    Const USER_PERMISSION_WU_IS_TSE As Integer = &H200000
    Const USER_PERMISSION_WU_IS_PILOT_USER As Integer = &H400000
    'Const USER_PERMISSION_PRODUCT_CREDITS_TAB As Integer = &H800000
    Const USER_PERMISSION_WU_IS_SALES As Integer = &H800000

    Const USER_PERMISSION_PRODUCT_CREDITS_TAB As Integer = &H1000000
    

    Const CUSTOMER_METHOD As Int32 = 713
    Const CUSTOMER_QUANTUM As Int32 = 774
    Const CUSTOMER_ARTHRITIS As Int32 = 711
    Const CUSTOMER_DEMO As Int32 = 16
    Const CUSTOMER_WURS As Int32 = 579
    Const CUSTOMER_WUIRE As Int32 = 686
    Const CUSTOMER_WURSDEMO As Int32 = 788
    Const CUSTOMER_WUFIN As Int32 = 798
    Const CUSTOMER_BIKES365 As Int32 = 749
    Const CUSTOMER_JUPITER As Int32 = 784
    Const CUSTOMER_POSITIVENOISE As Int32 = 821
    Const CUSTOMER_JAMESHARDIE As Int32 = 837
    Const CUSTOMER_HARDFR As Int32 = 849
    Const CUSTOMER_DECLAN As Int32 = 729
        
    Dim Tabx As VASPTBv4NET.ASPTabItem
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Dim propVisible As Boolean
    Dim propContinuousLoop As Boolean
    Dim propPauseOnMouseOver As Boolean
    Dim propScrollDirection As ComponentArt.Web.UI.ScrollDirection
    Dim propSmoothScrollSpeed As ComponentArt.Web.UI.SmoothScrollSpeed
    Dim propRotationType As ComponentArt.Web.UI.RotationType
    Dim propSlidePause As Integer
    Dim propScrollInterval As Integer
    Dim propShowEffect As ComponentArt.Web.UI.RotationEffect
    Dim propShowEffectDuration As Integer
    Dim propHideEffect As ComponentArt.Web.UI.RotationEffect
    Dim propHideEffectDuration As Integer
    
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not IsNumeric(Session("UserKey")) Then                               ' catch session expiry
            Server.Transfer("session_expired.aspx")
        Else
            Select Case Session("UserType")
                Case "User"
                    lblUserName.Text = Session("UserName") & " / "
                Case Else
                    lblUserName.Text = Session("UserName") & " (" & Session("UserType") & ") / "
            End Select
            lnkbtnCustomerName.Text = Session("CustomerName")
        End If
        'tbAccessKey.Attributes.Add("onkeypress", "return clickButton(event,'" + lnkbtnAccessKey.ClientID + "')")
        Call SetStyleSheet()
        If Session("LastLogon") <> DateTime.MinValue Then
            lblTopLineMessage.Text = "last login: " & Format(Session("LastLogon"), "dd-MMM-yy hh:mm")
            Session("LastLogon") = DateTime.MinValue
        End If
        If Session("RunningHeaderImage") = "default" Then
            imgRunningHeader.ImageUrl = ConfigLib.GetConfigItem_Default_Running_Header_Image
        ElseIf Session("RunningHeaderImage") <> "" Then
            imgRunningHeader.ImageUrl = Session("RunningHeaderImage")
        Else
            imgRunningHeader.Visible = False
        End If
    
        tdOldRotator.Visible = False
        tdNewRotator.Visible = False
        
        If UsesNewRotator() Then
            Call LoadControls()
            tdNewRotator.Visible = True
        Else
            InitFromXML()
            BindCustomerAdRotator()
            tdOldRotator.Visible = True
        End If
        
        TabView.CssClass = "TabView"
        If IsRAMBLERS() Then
            TabView.CssClass = "TabView_RAMBLERS"
        End If
        TabView.ImagePath = "./images"
        TabView.Orientation = 0
        TabView.TabBackColor = COLOUR_LIGHT_GREY
        TabView.SelectedBackColor = COLOUR_GREY   'grey
        TabView.SelectedForeColor = ""          'grey
        TabView.SelectedBold = True
        TabView.BodyBackground = COLOUR_WHITE
        TabView.TabWidth = 0
        TabView.StartTab = ""
        TabView.QueryString = ""
        TabView.LicenseKey = "D3FD-933B-6E8F"
        TabView.RunWithinSession = True
        Select Case Session("UserType")
            Case "sa", "SuperUser"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_VIEW_STOCK) > 0 Then
                    TabView.TabItem = ViewProductsTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_CREATE_STOCK_BOOKING) > 0 Then
                    If IsJupiter() Then
                        TabView.TabItem = OrdersTabJupiterStock()
                        TabView.TabItem = OrdersTabJupiterPOD()
                        'TabView.TabItem = SupplierUpdateTab()
                        TabView.TabItem = PrintStatusTab()
                        TabView.TabItem = FileUploadJUPITERTab()
                    Else
                        TabView.TabItem = OrdersTab()
                    End If
                    If IsPositiveNoise() Then
                        TabView.TabItem = PositiveNoiseOrderProcessorTab()
                    End If
                    'TabView.TabItem = OrdersTab()
                    If IsQuickOrder() Then
                        TabView.TabItem = QuickOrderTab()
                    End If
                End If
                If CBool(Session("AbleToCreateCollectionRequest")) = True Then
                    TabView.TabItem = CourierBookingsTab()
                    ' TabView.TabItem = TrackAndTraceAdvancedTab()
                End If
                TabView.TabItem = TrackAndTraceTab()
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_PRINT_ON_DEMAND_TAB) > 0 Then
                    TabView.TabItem = PrintOnDemandTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_PRODUCT_CREDITS_TAB) > 0 Then
                    TabView.TabItem = ProductCreditsTab()
                End If
                If IsMethod() Then
                    TabView.TabItem = FileUploadMETHODTab()
                End If
                If IsQuantum() Then
                    TabView.TabItem = QuantumAmazonOrderTab()
                End If
                If IsBikes365() Then
                    TabView.TabItem = AmazonOrderBikes365Tab()
                End If
                TabView.TabItem = ReportsTab()
                TabView.TabItem = AddressBookTab()
                TabView.TabItem = ProductManagerTab()

                If (IsWU() Or IsWURSDEMO()) Then
                    TabView.TabItem = UserManagerWUCustomTab()
                Else
                    TabView.TabItem = UserManagerTab()
                End If
                
                If Session("UserKey") = 5844 Then
                    TabView.TabItem = WUPermissionsTab()
                End If
                
                'If IsFEXCO() Then
                'If (CInt(Session("UserPermissions")) And USER_PERMISSION_ADVANCED_PERMISSIONS_TAB) > 0 Then
                '    TabView.TabItem = UserPermissionsTab()
                'End If
                TabView.TabItem = MyProfileTab()

                'If IsFEXCO() Then
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_FILE_UPLOAD_TAB) > 0 Then
                    Dim uriRequestUrl As Uri = Request.Url
                    If uriRequestUrl.Scheme = "https" Then
                        TabView.TabItem = FileUploadTab()
                    End If
                End If
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If

                If (IsWU() Or IsWURSDEMO()) Then
                    TabView.TabItem = HelpGuidesWUIRETab()
                End If

                If (Session("UserPermissions") And USER_PERMISSION_SITE_ADMINISTRATOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR) Then
                    TabView.TabItem = SiteAdministratorTab()
                End If
                If Session("UserPermissions") And USER_PERMISSION_ACCOUNT_HANDLER Then
                    TabView.TabItem = AccountHandlerTab()
                End If
            Case "Product Owner"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_VIEW_STOCK) > 0 Then
                    TabView.TabItem = ViewProductsTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_CREATE_STOCK_BOOKING) > 0 Then
                    If IsJupiter() Then
                        TabView.TabItem = OrdersTabJupiterStock()
                        'TabView.TabItem = OrdersTabJupiterPOD()
                    Else
                        TabView.TabItem = OrdersTab()
                    End If
                    'TabView.TabItem = OrdersTab()
                    If IsQuickOrder() Then
                        TabView.TabItem = QuickOrderTab()
                    End If
                End If
                If CBool(Session("AbleToCreateCollectionRequest")) = True Then
                    TabView.TabItem = CourierBookingsTab()
                    ' TabView.TabItem = TrackAndTraceAdvancedTab()
                End If
                TabView.TabItem = TrackAndTraceTab()
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_PRODUCT_CREDITS_TAB) > 0 Then
                    TabView.TabItem = ProductCreditsTab()
                End If
                TabView.TabItem = ReportsTab()
                TabView.TabItem = AddressBookTab()
                TabView.TabItem = ProductManagerTab()
                TabView.TabItem = MyProfileTab()
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If
                If (Session("UserPermissions") And USER_PERMISSION_SITE_ADMINISTRATOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_ADMINISTRATOR) Then
                    TabView.TabItem = SiteAdministratorTab()
                End If
            Case "User"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_VIEW_STOCK) > 0 Then
                    TabView.TabItem = ViewProductsTab()
                End If
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_CREATE_STOCK_BOOKING) > 0 Then
                    If Not IsWURSDEMO() Then
                        If IsJupiter() Then
                            TabView.TabItem = OrdersTabJupiterStock()
                            TabView.TabItem = OrdersTabJupiterPOD()
                        Else
                            TabView.TabItem = OrdersTab()
                        End If
                        'TabView.TabItem = OrdersTab()
                        If IsQuickOrder() Then
                            TabView.TabItem = QuickOrderTab()
                        End If
                        If IsPositiveNoise() Then
                            TabView.TabItem = PositiveNoiseOrderProcessorTab()
                        End If
                    Else
                        TabView.TabItem = QuickOrderWUTab()
                        TabView.TabItem = MessagingTab()
                    End If
                End If
                If CBool(Session("AbleToCreateCollectionRequest")) = True Then
                    TabView.TabItem = CourierBookingsTab()
                    'TabView.TabItem = TrackAndTraceAdvancedTab()
                End If
                TabView.TabItem = TrackAndTraceTab()
                If (CInt(Session("UserPermissions")) And USER_PERMISSION_PRODUCT_CREDITS_TAB) > 0 Then
                    TabView.TabItem = ProductCreditsTab()
                End If
                If Not (Session("UserPermissions") And USER_PERMISSION_VIEW_STOCK) > 0 Then
                    If Not IsWURSDEMO() Then
                        TabView.TabItem = AddressBookTab()
                    End If
                    'If IsLovells() Then
                    '    TabView.TabItem = PublicationManagerTab()
                    'End If
                    If Not (IsWU() Or IsWURSDEMO()) Then
                        TabView.TabItem = MyProfileTab()
                    End If
                End If
				If IsJamesHardie() Then
                    TabView.TabItem = JamesHardieReportTab()
				End If
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If
                    If (IsWU() Or IsWURSDEMO()) Then
                        TabView.TabItem = HelpGuidesWUIRETab()
                    End If
				Case "On Demand Supplier"
                TabView.TabItem = NoticeBoardTab()
                '    TabView.TabItem = OnDemandOrdersTab()
                '    If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                '        TabView.TabItem = SiteEditorTab()
                '    End If
                'TabView.TabItem = SupplierUpdateTab()
                TabView.TabItem = PrintStatusTab()
                ' TabView.TabItem = MyProfileTab()
            Case "Operations"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                TabView.TabItem = CustomersTab()
                TabView.TabItem = ProductReportGeneratorTab()
                TabView.TabItem = ConsumablesTab()
                TabView.TabItem = SerialNumbersTab()
                TabView.TabItem = WarehouseReportTab()
                TabView.TabItem = HGIAddresses()
                TabView.TabItem = ShelterOrderProcessorTab()
                TabView.TabItem = WUSerialNumbersTab()
                TabView.TabItem = WUIRESerialNumbersTab()
                TabView.TabItem = FEXCODeliveriesTab()
                TabView.TabItem = JobPricingTab()
                TabView.TabItem = TrackingCodesTab()
                'TabView.TabItem = ComparePermissionsTab()
                TabView.TabItem = ProductConfiguratorTab()
                TabView.TabItem = CloneProductsTab()
                TabView.TabItem = WayfairGoodsInTab()
                TabView.TabItem = WayfairStockAdjustmentTab()
                'TabView.TabItem = FEXCOStatementGeneratorTab()
                TabView.TabItem = HysterYaleOrdersTab()
                TabView.TabItem = OrderProcessorTab()
                TabView.TabItem = UploadProductsTab()
                TabView.TabItem = WarehouseLocationsEditorTab()
				TabView.TabItem = BlackRockStockAlertsByCostCentreTab()
                TabView.TabItem = HysterYaleInvoiceEditorTab()
                TabView.TabItem = QueryDBTab()
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If
            Case "PBOperations"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                TabView.TabItem = CustomersTab()
                TabView.TabItem = ProductReportGeneratorTab()
                TabView.TabItem = RyvitaOrderProcessorTab()
                TabView.TabItem = ConsumablesTab()
                TabView.TabItem = WarehouseReportTab()
                TabView.TabItem = JobPricingTab()
                TabView.TabItem = ProductReturnsTab()
                TabView.TabItem = GoodsInTab()
                TabView.TabItem = CloneProductsTab()
                TabView.TabItem = CloneConsignmentTab()
                TabView.TabItem = OrderProcessorTab()
                TabView.TabItem = UploadProductsTab()
                TabView.TabItem = WayfairGoodsInTab()
                TabView.TabItem = WayfairStockAdjustmentTab()
                TabView.TabItem = WarehouseLocationsEditorTab()
                TabView.TabItem = QueryDBTab()
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If
            Case "AllOperations"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                TabView.TabItem = CustomersTab()
                TabView.TabItem = ProductReportGeneratorTab()
                TabView.TabItem = ConsumablesTab()
                TabView.TabItem = SerialNumbersTab()
                TabView.TabItem = WarehouseReportTab()
                TabView.TabItem = HGIAddresses()
                TabView.TabItem = ShelterOrderProcessorTab()
                TabView.TabItem = WUSerialNumbersTab()
                TabView.TabItem = WUIRESerialNumbersTab()
                TabView.TabItem = FEXCODeliveriesTab()
                TabView.TabItem = JobPricingTab()
                TabView.TabItem = ProductReturnsTab()
                TabView.TabItem = GoodsInTab()
                TabView.TabItem = TrackingCodesTab()
                TabView.TabItem = ComparePermissionsTab()
                TabView.TabItem = ProductConfiguratorTab()
                TabView.TabItem = CloneProductsTab()
                TabView.TabItem = CloneConsignmentTab()
                TabView.TabItem = OrderProcessorTab()
                TabView.TabItem = UploadProductsTab()
                TabView.TabItem = WayfairDashboardTab()
                TabView.TabItem = WayfairGoodsInTab()
                TabView.TabItem = WayfairStockAdjustmentTab()
                TabView.TabItem = HysterYaleOrdersTab()
				TabView.TabItem = BlackRockStockAlertsByCostCentreTab()
                TabView.TabItem = WarehouseLocationsEditorTab()
                TabView.TabItem = FEXCOStatementGeneratorTab()
                TabView.TabItem = HysterYaleInvoiceEditorTab()
                TabView.TabItem = QueryDBTab()
                If (Session("UserPermissions") And USER_PERMISSION_SITE_EDITOR) Or (Session("UserPermissions") And USER_PERMISSION_DEPUTY_SITE_EDITOR) Then
                    If UsesNewRotator() Then
                        TabView.TabItem = SiteEditor2Tab()
                    Else
                        TabView.TabItem = SiteEditorTab()
                    End If
                End If
            Case "Wayfair"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                TabView.TabItem = WayfairDashboardTab()
                TabView.TabItem = WayfairGoodsInTab()
                TabView.TabItem = WayfairStockAdjustmentTab()
                TabView.TabItem = QueryDBTab()
                'Case "DataEntry"
                '    TabView.TabItem = NoticeBoardTab()
                '    TabView.TabItem = NHSMailingListTab()
            Case "WUOperations"
                If UsesNewRotator() Then
                    TabView.TabItem = NoticeBoard2Tab()
                Else
                    TabView.TabItem = NoticeBoardTab()
                End If
                TabView.TabItem = CustomersTab()
                TabView.TabItem = WUPrePaidCardsTab()
                TabView.TabItem = WUPermissionsTab()
                TabView.TabItem = WUSerialNumbersTab()
                TabView.TabItem = WUIRESerialNumbersTab()
                TabView.TabItem = FEXCODeliveriesTab()
                TabView.TabItem = WUFININTOrderTab()
                TabView.TabItem = UserManagerWUAgentAddressesTab()
                TabView.TabItem = WUFININTEditAgentTab()
                TabView.TabItem = WUCOSTAEditAgentTab()
				TabView.TabItem = WUMinGrabsTab()
                TabView.TabItem = WURSAgentAddressUploadTab()
				TabView.TabItem = VirtualProductsTab()
                TabView.TabItem = OrderProcessorTab()
                TabView.TabItem = FEXCOStatementGeneratorTab()
                ' TabView.TabItem = WUMICaptureTab()
                TabView.TabItem = QueryDBTab()
        End Select
    End Sub

    Protected Function UsesNewRotator() As Boolean
        If Session("CustomerCreatedOn") IsNot Nothing Then
            If CDate(Session("CustomerCreatedOn")) > DateTime.Parse("01-Aug-2014") Then
                Return True
            End If
        End If
        'Dim arrUsesNewRotator() As Integer = {CUSTOMER_WURS, CUSTOMER_WUIRE, CUSTOMER_WURSDEMO}
        Dim arrUsesNewRotator() As Integer = {CUSTOMER_WURSDEMO}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        UsesNewRotator = IIf(gbSiteTypeDefined, gsSiteType = "newrotator", Array.IndexOf(arrUsesNewRotator, nCustomerKey) >= 0)
    End Function

    Protected Function IsWU() As Boolean
        Dim arrWU() As Integer = {CUSTOMER_WURS, CUSTOMER_WUIRE, CUSTOMER_WUFIN, CUSTOMER_WURSDEMO}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWU = IIf(gbSiteTypeDefined, gsSiteType = "wu", Array.IndexOf(arrWU, nCustomerKey) >= 0)
    End Function

    Protected Function IsWURS() As Boolean
        Dim arrWURS() As Integer = {CUSTOMER_WURS}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsWURS = IIf(gbSiteTypeDefined, gsSiteType = "wurs", Array.IndexOf(arrWURS, nCustomerKey) >= 0)
    End Function

    Protected Function IsWURSDEMO() As Boolean
        IsWURSDEMO = False
        If CInt(Session("CustomerKey")) = CUSTOMER_WURSDEMO Then
            IsWURSDEMO = True
        End If
    End Function

    Protected Function IsMethod() As Boolean
        IsMethod = False
        If CInt(Session("CustomerKey")) = CUSTOMER_METHOD Then
            IsMethod = True
        End If
    End Function

    Protected Function IsJupiter() As Boolean
        IsJupiter = False
        If CInt(Session("CustomerKey")) = CUSTOMER_JUPITER Then
            IsJupiter = True
        End If
    End Function

    Protected Function IsPositiveNoise() As Boolean
        IsPositiveNoise = False
        If CInt(Session("CustomerKey")) = CUSTOMER_POSITIVENOISE Then
            IsPositiveNoise = True
        End If
    End Function

    Protected Function IsJamesHardie() As Boolean
        Dim arrCustomerJamesHardie() As Integer = {837, 849}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsJamesHardie = IIf(gbSiteTypeDefined, gsSiteType = "jameshardie", Array.IndexOf(arrCustomerJamesHardie, nCustomerKey) >= 0)
    End Function
	
    Protected Function IsQuantum() As Boolean
        IsQuantum = False
        If CInt(Session("CustomerKey")) = CUSTOMER_QUANTUM Then
            IsQuantum = True
        End If
    End Function

    Protected Function IsBikes365() As Boolean
        IsBikes365 = False
        If CInt(Session("CustomerKey")) = CUSTOMER_BIKES365 Then
            IsBikes365 = True
        End If
    End Function
	
    Protected Function IsRAMBLERS() As Boolean
        Dim arrWU() As Integer = {754}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsRAMBLERS = IIf(gbSiteTypeDefined, gsSiteType = "ramblers", Array.IndexOf(arrWU, nCustomerKey) >= 0)
    End Function

    Protected Function IsQuickOrder() As Boolean
        Dim arrQuickOrder() As Integer = {CUSTOMER_ARTHRITIS, CUSTOMER_DEMO, CUSTOMER_DECLAN}
        'Dim arrQuickOrder() As Integer = {CUSTOMER_ARTHRITIS}
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsQuickOrder = Array.IndexOf(arrQuickOrder, nCustomerKey) >= 0
    End Function

    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Function sVirtualAdrotatorThumbFolder() As String
        sVirtualAdrotatorThumbFolder = ConfigLib.GetConfigItem_Virtual_Thumb_URL
    End Function
    
    Protected Sub InitFromXML()
        Dim Doc As New XmlDataDocument
        Dim Nav As XPath.XPathNavigator
        Dim Iterator As XPath.XPathNodeIterator
        Dim Nav2 As XPath.XPathNavigator
        Dim sNodeName As String
        Dim sNodeValue As String
        
        Try
            Doc.Load(Server.MapPath(CONFIG_XML_FILE))                    ' Open the site configuration file
        Catch ex As System.IO.FileNotFoundException
            ' NewConfigFile
        End Try
    
        Nav = CType(Doc, XPath.IXPathNavigable).CreateNavigator()         ' Set nav object
        Try
            Iterator = Nav.Select("DisplaySettings/HeaderRotator/*")      ' Set node iterator 
            While Iterator.MoveNext()                                     ' Move to the desired node
                Nav2 = Iterator.Current.Clone()                           ' Get the value of the current node
                sNodeName = Nav2.LocalName
                sNodeValue = Iterator.Current.Value
    
                Select Case sNodeName
    
                    Case "Visible"
                        If sNodeValue = "True" Then
                            propVisible = True
                        ElseIf sNodeValue = "False" Then
                            propVisible = False
                        Else : propVisible = True
                        End If
    
                    Case "ContinuousLoop"
                        If sNodeValue = "True" Then
                            propContinuousLoop = True
                        ElseIf sNodeValue = "False" Then
                            propContinuousLoop = False
                        End If
    
                    Case "PauseOnMouseOver"
                        If sNodeValue = "True" Then
                            propPauseOnMouseOver = True
                        ElseIf sNodeValue = "False" Then
                            propPauseOnMouseOver = False
                        End If
    
                    Case "ScrollDirection"
                        If sNodeValue = "Up" Then
                            propScrollDirection = ComponentArt.Web.UI.ScrollDirection.Up
                        ElseIf sNodeValue = "Left" Then
                            propScrollDirection = ComponentArt.Web.UI.ScrollDirection.Left
                        End If
    
                    Case "SmoothScrollSpeed"
                        If sNodeValue = "Slow" Then
                            propSmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Slow
                        ElseIf sNodeValue = "Medium" Then
                            propSmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Medium
                        ElseIf sNodeValue = "Fast" Then
                            propSmoothScrollSpeed = ComponentArt.Web.UI.SmoothScrollSpeed.Fast
                        End If
    
                    Case "RotationType"
                        If sNodeValue = "ContentScroll" Then
                            propRotationType = ComponentArt.Web.UI.RotationType.ContentScroll
                        ElseIf sNodeValue = "SlideShow" Then
                            propRotationType = ComponentArt.Web.UI.RotationType.SlideShow
                        End If
    
                    Case "SlidePause"
                        propSlidePause = CInt(sNodeValue)
    
                    Case "ScrollInterval"
                        propScrollInterval = CInt(sNodeValue)
    
                    Case "ShowEffect"
                        If sNodeValue = "None" Then
                            propShowEffect = ComponentArt.Web.UI.RotationEffect.None
                        ElseIf sNodeValue = "Fade" Then
                            propShowEffect = ComponentArt.Web.UI.RotationEffect.Fade
                        ElseIf sNodeValue = "Pixelate" Then
                            propShowEffect = ComponentArt.Web.UI.RotationEffect.Pixelate
                        ElseIf sNodeValue = "Dissolve" Then
                            propShowEffect = ComponentArt.Web.UI.RotationEffect.Dissolve
                        ElseIf sNodeValue = "GradientWipe" Then
                            propShowEffect = ComponentArt.Web.UI.RotationEffect.GradientWipe
                        End If
    
                    Case "ShowEffectDuration"
                        propShowEffectDuration = CInt(sNodeValue)
    
                    Case "HideEffect"
                        If sNodeValue = "None" Then
                            propHideEffect = ComponentArt.Web.UI.RotationEffect.None
                        ElseIf sNodeValue = "Fade" Then
                            propHideEffect = ComponentArt.Web.UI.RotationEffect.Fade
                        ElseIf sNodeValue = "Pixelate" Then
                            propHideEffect = ComponentArt.Web.UI.RotationEffect.Pixelate
                        ElseIf sNodeValue = "Dissolve" Then
                            propHideEffect = ComponentArt.Web.UI.RotationEffect.Dissolve
                        ElseIf sNodeValue = "GradientWipe" Then
                            propHideEffect = ComponentArt.Web.UI.RotationEffect.GradientWipe
                        End If
    
                    Case "HideEffectDuration"
                        propHideEffectDuration = CInt(sNodeValue)
    
                    Case Else
    
                End Select
    
            End While
    
        Catch ex As System.Xml.XPath.XPathException
            'lblMessage.Text = "Bad XPath query"
        End Try
    End Sub

    Private Sub LoadControls()
        
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetPagePanelControlsFromCustomerKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        
        Try
            oAdapter.Fill(oDataTable)
            lblRotator.Text = String.Empty
            If oDataTable.Rows.Count = 1 Then
                Dim dr As DataRow = oDataTable.Rows(0)
                lblRotator.Text = dr("Header")
                FindControlByTags(lblRotator)
            Else
                If oDataTable.Rows.Count > 1 Then
                    WebMsgBox.Show("LoadControls: More than one header row found.")
                End If
            End If
        Catch ex As Exception
            WebMsgBox.Show("GetPageContent: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
    End Sub
    
    Protected Function FindControlByTags(ByVal lblHTML As Label) As String
        
        'Dim regularExpressionPattern As String = "\[(.*?)\]"
        Dim regularExpressionPattern As String = "\[[a-zA-Z0-9]*.*\]"
        Dim re As New Regex(regularExpressionPattern)
        
        For Each oMatch In re.Matches(lblHTML.Text)
            
            CreateDynamicControl(oMatch.ToString(), lblHTML)
            
        Next
        
        Return lblHTML.Text
        
    End Function
    
    Protected Sub CreateDynamicControl(ByVal controlTags As String, ByVal ctllbl As Label)
        
        Dim sDataSourceTag As String = String.Empty
        
        If controlTags.ToLower.Contains("rotator") Then
            
            Dim strTemplateControls As String() = controlTags.Split(":")
            Dim rr As New RadRotator
            
            If strTemplateControls.Length > 0 Then
                rr.ItemTemplate = New RadRotatorTemplate(strTemplateControls)
            Else
                rr.ItemTemplate = New RadRotatorTemplate
            End If
            
            
            If controlTags.ToLower.Contains(";") Then
                Dim sRotatorProperties As String() = controlTags.Split(";")
                For Each config As String In sRotatorProperties
                    If config.ToLower().Contains("height=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.Height = Convert.ToInt32(sValue)
                        Else
                            rr.Height = 100
                        End If
                    ElseIf config.ToLower().Contains("width=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.Width = Convert.ToInt32(sValue)
                        Else
                            rr.Width = 400
                        End If
                    ElseIf config.ToLower().Contains("scrolldirection=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "up" Then
                                rr.ScrollDirection = RotatorScrollDirection.Up
                            ElseIf sValue.ToLower() = "down" Then
                                rr.ScrollDirection = RotatorScrollDirection.Down
                            ElseIf sValue.ToLower() = "left" Then
                                rr.ScrollDirection = RotatorScrollDirection.Left
                            ElseIf sValue.ToLower() = "right" Then
                                rr.ScrollDirection = RotatorScrollDirection.Right
                            Else
                                rr.ScrollDirection = RotatorScrollDirection.Up
                            End If
                        End If
                    ElseIf config.ToLower().Contains("scrollduration=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If IsNumeric(sValue) Then
                            rr.ScrollDuration = Convert.ToInt32(sValue)
                        Else
                            rr.ScrollDuration = 3000
                        End If
                    ElseIf config.ToLower().Contains("datasource") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            sDataSourceTag = sValue
                        End If
                    ElseIf config.ToLower.Contains("rotatortype=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "buttons" Then
                                rr.RotatorType = RotatorType.Buttons
                            ElseIf sValue.ToLower() = "automaticadvance" Then
                                rr.RotatorType = RotatorType.AutomaticAdvance
                            ElseIf sValue.ToLower() = "buttonsover" Then
                                rr.RotatorType = RotatorType.ButtonsOver
                            ElseIf sValue.ToLower() = "carousel" Then
                                rr.RotatorType = RotatorType.Carousel
                            ElseIf sValue.ToLower() = "carouselbuttons" Then
                                rr.RotatorType = RotatorType.CarouselButtons
                            ElseIf sValue.ToLower() = "coverflow" Then
                                rr.RotatorType = RotatorType.CoverFlow
                            ElseIf sValue.ToLower() = "coverflowbuttons" Then
                                rr.RotatorType = RotatorType.CoverFlowButtons
                            ElseIf sValue.ToLower() = "slideshow" Then
                                rr.RotatorType = RotatorType.SlideShow
                            ElseIf sValue.ToLower() = "slideshowbuttons" Then
                                rr.RotatorType = RotatorType.SlideShowButtons
                            Else
                                rr.RotatorType = RotatorType.Carousel
                            End If
                        End If
                    ElseIf config.ToLower.Contains("animationtype=") Then
                        Dim sconfigValue() As String = config.Split("=")
                        Dim sValue As String = sconfigValue(1).Replace("]", String.Empty)
                        If Not String.IsNullOrEmpty(sValue) Then
                            If sValue.ToLower() = "none" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.None
                            ElseIf sValue.ToLower() = "fade" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Fade
                            ElseIf sValue.ToLower() = "pulse" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.Pulse
                            ElseIf sValue.ToLower() = "crossfade" Then
                                rr.SlideShowAnimation.Type = Telerik.Web.UI.Rotator.AnimationType.CrossFade
                            End If
                        End If
                    End If
                Next
            End If

            'rr.RotatorType = RotatorType.Buttons
            'rr.ScrollDirection = RotatorScrollDirection.Up
            'rr.ScrollDuration = 3000
            'rr.Height = 100
            'rr.Width = 400
            rr.ItemHeight = 50
            rr.DataSource = BindRotator(sDataSourceTag)
            rr.DataBind()
            
            ctllbl.Controls.Add(rr)
            
        End If
        
        If controlTags.ToLower.Contains("textbox:") Then
            
            Dim stringSeparators() As String = {"[textbox:"}
            Dim strTemplateControls As String() = controlTags.Split(stringSeparators, StringSplitOptions.None)
            If strTemplateControls.Length > 1 Then
                
                Dim lbl As New Label
                lbl.Text = strTemplateControls(1).ToString.Replace("]", "")
                ctllbl.Controls.Add(lbl)
            End If
            
        End If
        
    End Sub
    
    Protected Function BindRotator(ByVal sDataSourceTag As String) As DataTable
        
        BindRotator = Nothing
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_RotatorGetContentFromCustomerKey", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Convert.ToInt32(Session("CustomerKey"))
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@DataSourceTag", SqlDbType.VarChar))
        oAdapter.SelectCommand.Parameters("@DataSourceTag").Value = sDataSourceTag
        
        
        Try
            oAdapter.Fill(oDataTable)
            BindRotator = oDataTable
            
        Catch ex As Exception
            WebMsgBox.Show("GetPageContent: " & ex.Message)
            
        Finally
            oConn.Close()
        End Try
        
    End Function
    
    Protected Function ExtractFromRSSFeed() As DataTable
        
        Dim dt As DataTable = CreateProductsDataTable()
        Dim dr As DataRow
        Dim nRssCount As Integer = 0
        Dim sRssUrl As String = String.Empty
        Dim nCustomerKey As Integer = Convert.ToInt64(Session("CustomerKey"))
        If nCustomerKey > 0 Then
            Dim sQuery As String = "select RssUrl, RssCount from RotatorRssFeed where CustomerKey = " & nCustomerKey
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
            If oDataTable IsNot Nothing AndAlso oDataTable.Rows.Count <> 0 Then
                dr = oDataTable.Rows(0)
                If Not IsDBNull(dr("RssCount")) Then
                    nRssCount = Convert.ToInt32(dr("RssCount"))
                End If
                If Not IsDBNull(dr("RssUrl")) Then
                    sRssUrl = dr("RssUrl").ToString()
                End If
            End If
        End If
        
        
        Dim sf As SyndicationFeed = LoadRSSFeed(sRssUrl)
        If Not sf Is Nothing Then
            For Each item As SyndicationItem In sf.Items
                'Dim item As SyndicationItem = sf.Items(i)
                dr = dt.NewRow
                dr("Title") = item.Title.Text
                dr("Date") = Convert.ToDateTime(item.PublishDate.LocalDateTime).ToString("dd-MMM-yyyy HH:mm")
                dr("Content") = item.Summary.Text
                If Not item.Links Is Nothing Then
                    For Each sl As SyndicationLink In item.Links
                        dr("BaseURI") = sl.Uri.AbsoluteUri
                    Next
                End If
                Dim bFirst As Boolean = True
                For Each sc As SyndicationCategory In item.Categories
                    If Not bFirst Then
                        dr("Categories") += ", "
                    End If
                    'dr("Categories") += sc.Name
                    dr("Categories") += sc.Name
                    bFirst = False
                Next
                dt.Rows.Add(dr)
                'If i = CInt(nRssCount) - 1 Then
                '    Exit For
                'End If
            Next
        End If
        ExtractFromRSSFeed = dt
    End Function
    
    Protected Function LoadRSSFeed(ByVal sRssUrl As String) As SyndicationFeed
        Try
            Using reader As XmlReader = XmlReader.Create(sRssUrl)
                LoadRSSFeed = SyndicationFeed.Load(reader)
            End Using
        Catch ex As WebException
        Catch ex As XmlException
        Catch ex As Exception
        End Try
    End Function
    
    Protected Function CreateProductsDataTable() As DataTable
        Dim oDataTable As New DataTable("RotatorRssFeed")
        oDataTable.Columns.Add(New DataColumn("Title", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Content", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("ImageTag", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Date", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("BaseUri", GetType(String)))
        oDataTable.Columns.Add(New DataColumn("Categories", GetType(String)))
        CreateProductsDataTable = oDataTable
    End Function
    
    Protected Sub BindCustomerAdRotator()
        Dim sConn As String = ConfigLib.GetConfigItem_ConnectionString
        Dim oConn As New SqlConnection(sConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_GetRotatorAds", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        Try
            oAdapter.Fill(oDataSet, "Ads")
            Dim Source As DataView = oDataSet.Tables("Ads").DefaultView
            Rotator1.Visible = False
            If Source.Count > 0 And blnVisible = True Then
                Rotator1.DataSource = Source
                Rotator1.DataBind()
                Rotator1.Visible = True
            Else
                Rotator1.Visible = False
            End If
        Catch ex As SqlException
            'lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub btn_LogOut_click(ByVal s As Object, ByVal e As EventArgs)
        Dim sLogonPage As String
        Call RemoveAutoLogonCookie()
        sLogonPage = ConfigLib.GetConfigItem_Default_Logon_Page
        Session.Abandon()
        Server.Transfer(sLogonPage)
    End Sub

    Protected Sub RemoveAutoLogonCookie()
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
    
    Protected Function JamesHardieReportTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "jhreport"
        Tabx.Text = "Order History"
        Tabx.URL = "reports_hardieuser.aspx"
        Tabx.ToolTipText = "order history"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function HysterYaleInvoiceEditorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "hyinvoiceeditor"
        Tabx.Text = "Hyster/Yale Invoice Editor"
        Tabx.URL = "hysteryaleinvoiceeditor.aspx"
        Tabx.ToolTipText = "hyster/yale invoice editor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
	Protected Function MessagingTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "messaging"
        Tabx.Text = "Agent Support"
        Tabx.URL = "Messaging.aspx"
        Tabx.ToolTipText = "messaging"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function CloneProductsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "cloneproducts"
        Tabx.Text = "Clone Products"
        Tabx.URL = "CloneProducts.aspx"
        Tabx.ToolTipText = "clone products"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function NoticeBoardTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "noticeboard"
        Tabx.Text = "Notice Board"
        Tabx.URL = "NoticeBoard.aspx"
        Tabx.ToolTipText = "notice board"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function NoticeBoard2Tab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "noticeboard2"
        Tabx.Text = "Notice Board"
        Tabx.URL = "NoticeBoard2.aspx"
        Tabx.ToolTipText = "notice board"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function CourierBookingsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "collections"
        Tabx.Text = "Courier Bookings"
        Tabx.URL = "courier_collection.aspx"
        Tabx.ToolTipText = "courier pickup requests"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function OrdersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "orders"
        Tabx.Text = "Place an Order"
        Tabx.URL = "on_line_picks.aspx"
        Tabx.ToolTipText = "place an order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function OrdersTabJupiterStock() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "ordersjupstock"
        Tabx.Text = "Stock Orders"
        Tabx.URL = "on_line_picks.aspx"
        Tabx.ToolTipText = "place a stock order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function VirtualProductsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "virtualproducts"
        Tabx.Text = "Virtual Products"
        Tabx.URL = "ProductManagerVirtual.aspx"
        Tabx.ToolTipText = "Manage Virtual Products"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function OrdersTabJupiterPOD() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "ordersjuppod"
        Tabx.Text = "Print Orders"
        Tabx.URL = "on_line_picks_juppod.aspx"
        Tabx.ToolTipText = "place a print order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function QuickOrderTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "quickorder"
        Tabx.Text = "Quick Order"
        Tabx.URL = "quickorder.aspx"
        Tabx.ToolTipText = "place an order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function QuickOrderWUTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "quickorderwu"
        Tabx.Text = "Place an Order"
        Tabx.URL = "quickorderwu.aspx"
        Tabx.ToolTipText = "place an order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function HysterYaleOrdersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "hysteryaleorders"
        Tabx.Text = "Hyster / Yale Orders"
        Tabx.URL = "HysterYaleOrders.aspx"
        Tabx.ToolTipText = "Hyster / Yale orders"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ViewProductsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "viewproducts"
        Tabx.Text = "View Products"
        Tabx.URL = "on_line_picks.aspx"
        Tabx.ToolTipText = "view products"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ConsumablesTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "consumables"
        Tabx.Text = "Consumables"
        Tabx.URL = "consumables.aspx"
        Tabx.ToolTipText = "consumables"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function TrackAndTraceTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx = New VASPTBv4NET.ASPTabItem()
        Tabx.Key = "trackandtrace"
        Tabx.Text = "Track & Trace"
        Tabx.URL = "track_and_trace.aspx"
        Tabx.ToolTipText = "track your consignments"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function TrackAndTraceAdvancedTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "trackandtraceadv"
        Tabx.Text = "Advanced Tracking"
        Tabx.URL = "advancedtracking.aspx"
        Tabx.ToolTipText = "track your consignments"
        Call SetCommonTabProperties(Tabx)
        TabView.TabItem = Tabx
        Return Tabx
    End Function

    Protected Function ProductCreditsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "productcredits"
        Tabx.Text = "Authorisations"
        Tabx.URL = "ProductCreditOverdrafts.aspx"
        Tabx.ToolTipText = "authorisation requests"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function ReportsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "reports"
        Tabx.Text = "Reports"
        Tabx.URL = "reports.aspx"
        Tabx.ToolTipText = "reports"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ProductConfiguratorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "productconfigurator"
        Tabx.Text = "Product Configurator"
        Tabx.URL = "ProductConfigurator.aspx"
        Tabx.ToolTipText = "product configurator"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function AddressBookTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "addressbook"
        Tabx.Text = "Address Book"
        Tabx.URL = "address_book.aspx"
        Tabx.ToolTipText = "manage your address book"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ProductManagerTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "productmanager"
        Tabx.Text = "Product Manager"
        Tabx.URL = "product_manager.aspx"
        Tabx.ToolTipText = "add / edit products"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    'Protected Function PublicationManagerTab() As VASPTBv4NET.ASPTabItem
    '    Dim Tabx As New VASPTBv4NET.ASPTabItem
    '    Tabx.Key = "publicationmanager"
    '    Tabx.Text = "Publication Manager"
    '    Tabx.URL = "PublicationManager.aspx"
    '    Tabx.ToolTipText = "add / edit publications"
    '    Call SetCommonTabProperties(Tabx)
    '    Return Tabx
    'End Function
    
    Protected Function UserManagerTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "users"
        Tabx.Text = "User Manager"
        Tabx.URL = "user_manager.aspx"
        Tabx.ToolTipText = "add / edit users"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function FileUploadTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "fileupload"
        Tabx.Text = "File Upload"
        Tabx.URL = "FileUpload.aspx"
        Tabx.ToolTipText = "secure file upload"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function FileUploadMETHODTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "fileuploadMETHOD"
        Tabx.Text = "File Upload"
        Tabx.URL = "FileUploadMETHOD.aspx"
        Tabx.ToolTipText = "file upload"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function FileUploadJUPITERTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "fileuploadJUPITER"
        Tabx.Text = "File Upload"
        Tabx.URL = "FileUploadJUPITER.aspx"
        Tabx.ToolTipText = "file upload"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function QuantumAmazonOrderTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "AmazonOrder"
        Tabx.Text = "Amazon"
        Tabx.URL = "QuantumProcessAmazonOrder.aspx"
        Tabx.ToolTipText = "Process Amazon order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function AmazonOrderBikes365Tab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "AmazonOrderBikes365"
        Tabx.Text = "Amazon"
        Tabx.URL = "ProcessAmazonOrderBikes365.aspx"
        Tabx.ToolTipText = "Process Amazon order"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function AccountHandlerTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "accounthandler"
        Tabx.Text = "Account Handler"
        Tabx.URL = "AccountHandler.aspx"
        Tabx.ToolTipText = "account handler"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function SiteAdministratorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "siteadministrator"
        Tabx.Text = "Site Administrator"
        Tabx.URL = "Administrator.aspx"
        Tabx.ToolTipText = "administrator"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function SiteEditorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "siteeditor"
        Tabx.Text = "Site Editor"
        Tabx.URL = "SiteEditor.aspx"
        Tabx.ToolTipText = "edit the login and notice board page"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function SiteEditor2Tab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "siteeditor2"
        Tabx.Text = "Site Editor"
        Tabx.URL = "SiteEditor2.aspx"
        Tabx.ToolTipText = "edit the login and notice board page"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ProjectsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "projects"
        Tabx.Text = "Projects"
        Tabx.URL = "projects.aspx"
        Tabx.ToolTipText = "projects"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function MyProfileTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "myprofile"
        Tabx.Text = "My Profile"
        Tabx.URL = "MyProfile.aspx"
        Tabx.ToolTipText = "my profile"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function PrintOnDemandTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "printondemand"
        Tabx.Text = "Print On Demand"
        Tabx.URL = "PrintOnDemand.aspx"
        Tabx.ToolTipText = "print on demand"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function SupplierUpdateTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "supplierupdate"
        Tabx.Text = "Supplier Update"
        Tabx.URL = "SupplierUpdate.aspx"
        Tabx.ToolTipText = "supplier update"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function PrintStatusTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "printstatus"
        Tabx.Text = "Print Status"
        Tabx.URL = "PrintStatus.aspx"
        Tabx.ToolTipText = "print status"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function OnDemandOrdersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "ondemandorders"
        Tabx.Text = "On Demand Orders"
        Tabx.URL = "OnDemandOrders.aspx"
        Tabx.ToolTipText = "on demand orders"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function SerialNumbersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "manserialnumbers"
        Tabx.Text = "MAN Serial Nos"
        Tabx.URL = "SerialNumbers.aspx"
        Tabx.ToolTipText = "MAN serial numbers"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function KodakStorageTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "kodakstorage"
        Tabx.Text = "Kodak Storage"
        Tabx.URL = "KodakStorage.aspx"
        Tabx.ToolTipText = "Kodak monthly storage"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function FEXCOStorageTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wustorage"
        Tabx.Text = "WU Storage"
        Tabx.URL = "FEXCOStorage.aspx"
        Tabx.ToolTipText = "WU (UK) monthly storage"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function WUSerialNumbersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wuserialnumbers"
        Tabx.Text = "WU (UK) Serial Nos"
        Tabx.URL = "FEXCOSerialNumbers.aspx"
        Tabx.ToolTipText = "WU (UK) serial nos"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function WUIRESerialNumbersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wuireserialnumbers"
        Tabx.Text = "WUIRE Serial Nos"
        Tabx.URL = "WUIRESerialNumbers.aspx"
        Tabx.ToolTipText = "WUIRE serial nos"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function FEXCODeliveriesTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wudeliveries"
        Tabx.Text = "WU (UK) Goods In (Delivs)"
        Tabx.URL = "FEXCODeliveries.aspx"
        Tabx.ToolTipText = "WU goods in"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function ExtractPalletCountTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "extractpalletcount"
        Tabx.Text = "Extract Pallet Count"
        Tabx.URL = "PalletCountExtract.aspx"
        Tabx.ToolTipText = "Extract pallet count"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function FEXCOStatementGeneratorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUstatementgenerator"
        Tabx.Text = "WU Statements"
        Tabx.URL = "FEXCOStatementGenerator.aspx"
        Tabx.ToolTipText = "WU statement generator"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function QueryDBTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "querydb"
        Tabx.Text = "Database Query"
        Tabx.URL = "QueryDB.aspx"
        Tabx.ToolTipText = "query the database"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function UserPermissionsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "userpermissions"
        Tabx.Text = "User Permissions"
        Tabx.URL = "UserPermissions.aspx"
        Tabx.ToolTipText = "user permissions"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function ComparePermissionsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "comparepermissions"
        Tabx.Text = "Compare Perms"
        Tabx.URL = "ComparePermissions.aspx"
        Tabx.ToolTipText = "compare permissions"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function CustomersTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "customers"
        Tabx.Text = "Customers"
        Tabx.URL = "Customers.aspx"
        Tabx.ToolTipText = "customers"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function NHSMailingListTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "NHSMailingList"
        Tabx.Text = "NHS Mailing List"
        Tabx.URL = "NHSMailingList.aspx"
        Tabx.ToolTipText = "NHS Mailing List"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
    
    Protected Function WarehouseReportTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "warehouse"
        Tabx.Text = "Warehouse"
        Tabx.URL = "WarehouseReport.aspx"
        Tabx.ToolTipText = "Warehouse Report"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function WayfairDashboardTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wayfairdashboard"
        Tabx.Text = "Wayfair Dashboard"
        Tabx.URL = "WayfairDashboard.aspx"
        Tabx.ToolTipText = "Wayfair Dashboard"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function WayfairGoodsInTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wayfairgoodsin"
        Tabx.Text = "Wayfair Goods In"
        Tabx.URL = "WayfairGoodsIn.aspx"
        Tabx.ToolTipText = "Wayfair Goods In"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function WayfairStockAdjustmentTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "wayfairstockadjustments"
        Tabx.Text = "Wayfair Stock Adjustments"
        Tabx.URL = "WayfairInventoryAdjustments.aspx"
        Tabx.ToolTipText = "Wayfair Stock Adjustments"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function JobPricingTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "jobpricing"
        Tabx.Text = "Job Pricing"
        Tabx.URL = "JobPricing.aspx"
        Tabx.ToolTipText = "peterborough job pricing"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ProductReturnsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "productreturns"
        Tabx.Text = "Returns"
        Tabx.URL = "ProductReturns.aspx"
        Tabx.ToolTipText = "product returns"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function GoodsInTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "goodsin"
        Tabx.Text = "Goods In"
        Tabx.URL = "GoodsIn.aspx"
        Tabx.ToolTipText = "goods in"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function TrackingCodesTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "trackingcodes"
        Tabx.Text = "Tracking Codes"
        Tabx.URL = "TrackingCodes.aspx"
        Tabx.ToolTipText = "edit tracking codes"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function HGIAddresses() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "hgiaddresses"
        Tabx.Text = "HGI Addresses"
        Tabx.URL = "HGIAddresses.aspx"
        Tabx.ToolTipText = "HGI Addresses"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ShelterOrderProcessorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "shelterorderprocessor"
        Tabx.Text = "Shelter Order Processor"
        Tabx.URL = "ShelterOrderProcessor.aspx"
        Tabx.ToolTipText = "Shelter Order Processor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function

    Protected Function ProductReportGeneratorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "ProductReportGenerator"
        Tabx.Text = "Product Report Generator"
        Tabx.URL = "ProductReportGenerator.aspx"
        Tabx.ToolTipText = "Product Report Generator"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    'Protected Function WUPermissioningTab() As VASPTBv4NET.ASPTabItem
    '    Dim Tabx As New VASPTBv4NET.ASPTabItem
    '    Tabx.Key = "WUPermissioning"
    '    Tabx.Text = "WU Permissioning"
    '    Tabx.URL = "WUPermissioning.aspx"
    '    Tabx.ToolTipText = "Western Union Permissioning"
    '    Call SetCommonTabProperties(Tabx)
    '    Return Tabx
    'End Function
	
    Protected Function WUPermissionsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUPermissions"
        Tabx.Text = "WU Permissions"
        Tabx.URL = "WUPermissions.aspx"
        Tabx.ToolTipText = "Western Union Permissions"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function UserManagerWUCustomTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "UserManagerWUCustom"
        Tabx.Text = "WU Custom User Manager"
        Tabx.URL = "UserManagerWUCustom.aspx"
        Tabx.ToolTipText = "Only for use by Marilyn Quinn managing WU"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function UserManagerWUAgentAddressesTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "UserManagerWUAgentAddresses"
        Tabx.Text = "Edit WURS/WUIRE Agent"
        Tabx.URL = "WUAddEditAgents.aspx"
        Tabx.ToolTipText = "Add / edit WU agent addresses"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	    
    Protected Function CloneConsignmentTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "CloneConsignment"
        Tabx.Text = "Clone Consignment"
        Tabx.URL = "CloneConsignment.aspx"
        Tabx.ToolTipText = "Clone Consignment"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUFININTOrderTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUFININTOrder"
        Tabx.Text = "FININT / COSTA Order"
        Tabx.URL = "WUFININTOrder.aspx"
        Tabx.ToolTipText = "Process FININT or COSTA Order File"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUFININTEditAgentTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUFININTEditAgent"
        Tabx.Text = "Edit FININT Agent"
        Tabx.URL = "WUFININTEditAgent.aspx"
        Tabx.ToolTipText = "Edit FININT Agent Details"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUCOSTAEditAgentTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUCOSTAEditAgent"
        Tabx.Text = "Edit COSTA Agent"
        Tabx.URL = "WUCOSTAEditAgent.aspx"
        Tabx.ToolTipText = "Edit COSTA Agent Details"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WURSAgentAddressUploadTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WURSAgentAddressUpload"
        Tabx.Text = "WURS Agent Address Upload"
        Tabx.URL = "WURSAgentAddressUpload.aspx"
        Tabx.ToolTipText = "WURS Agent Address Upload"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUMinGrabsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUMinGrabs"
        Tabx.Text = "WU Min Grabs"
        Tabx.URL = "WUMinGrabs.aspx"
        Tabx.ToolTipText = "WU Min Grabs"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUMICaptureTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "WUMICapture"
        Tabx.Text = "Capture MI Data"
        Tabx.URL = "WUCaptureMIMonthlyData.aspx"
        Tabx.ToolTipText = "Capture MI Data"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function RyvitaOrderProcessorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "RyvitaOrderProcessor"
        Tabx.Text = "Ryvita Order Processor"
        Tabx.URL = "RyvitaOrderProcessor.aspx"
        Tabx.ToolTipText = "Ryvita Order Processor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WUPrePaidCardsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "PrepaidCards"
        Tabx.Text = "Prepaid Cards"
        Tabx.URL = "WUPrepaidCards.aspx"
        Tabx.ToolTipText = "Prepaid Cards audit trail"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function OrderProcessorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "orderprocessor"
        Tabx.Text = "Order Processor"
        Tabx.URL = "OrderProcessor.aspx"
        Tabx.ToolTipText = "Order Processor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function PositiveNoiseOrderProcessorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "positivenoiseorderprocessor"
        Tabx.Text = "Order Processor"
        Tabx.URL = "PositiveNoiseOrderProcessor.aspx"
        Tabx.ToolTipText = "Order Processor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function UploadProductsTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "uploadproducts"
        Tabx.Text = "Upload Products"
        Tabx.URL = "UploadProducts.aspx"
        Tabx.ToolTipText = "Upload Products"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function WarehouseLocationsEditorTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "warehouselocationseditor"
        Tabx.Text = "WH Locations"
        Tabx.URL = "WarehouseLocationsEditor.aspx"
        Tabx.ToolTipText = "Warehouse Locations Editor"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function BlackRockStockAlertsByCostCentreTab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "blockrockcostcentre"
        Tabx.Text = "BlackRock CCs"
        Tabx.URL = "BlackRockStockAlertsByCostCentre.aspx"
        Tabx.ToolTipText = "BlackRock Stock Alerts By Cost Centre"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Function HelpGuidesWUIRETab() As VASPTBv4NET.ASPTabItem
        Dim Tabx As New VASPTBv4NET.ASPTabItem
        Tabx.Key = "helpguideswuire"
        Tabx.Text = "Help Guides"
        Tabx.URL = "HelpGuidesWUIRE.aspx"
        Tabx.ToolTipText = "Help Guides"
        Call SetCommonTabProperties(Tabx)
        Return Tabx
    End Function
	
    Protected Sub SetCommonTabProperties(ByRef Tabx As VASPTBv4NET.ASPTabItem)
        Tabx.Image = "arrow.gif"
        Tabx.ForceDHTML = False
        Tabx.DHTML = ""
    End Sub

    ReadOnly Property blnVisible() As Boolean
        Get
            Return propVisible
        End Get
    End Property
    
    ReadOnly Property blnContinuousLoop() As Boolean
        Get
            Return propContinuousLoop
        End Get
    End Property
    
    ReadOnly Property blnPauseOnMouseOver() As Boolean
        Get
            Return propPauseOnMouseOver
        End Get
    End Property
    
    ReadOnly Property enumScrollDirection() As ComponentArt.Web.UI.ScrollDirection
        Get
            Return propScrollDirection
        End Get
    End Property
    
    ReadOnly Property enumSmoothScrollSpeed() As ComponentArt.Web.UI.SmoothScrollSpeed
        Get
            Return propSmoothScrollSpeed
        End Get
    End Property
    
    ReadOnly Property enumRotationType() As ComponentArt.Web.UI.RotationType
        Get
            Return propRotationType
        End Get
    End Property
    
    ReadOnly Property intSlidePause() As Integer
        Get
            Return propSlidePause
        End Get
    End Property
    
    ReadOnly Property intScrollInterval() As Integer
        Get
            Return propScrollInterval
        End Get
    End Property
    
    ReadOnly Property enumShowEffect() As ComponentArt.Web.UI.RotationEffect
        Get
            Return propShowEffect
        End Get
    End Property
    
    ReadOnly Property intShowEffectDuration() As Integer
        Get
            Return propShowEffectDuration
        End Get
    End Property
    
    ReadOnly Property enumHideEffect() As ComponentArt.Web.UI.RotationEffect
        Get
            Return propHideEffect
        End Get
    End Property
    
    ReadOnly Property intHideEffectDuration() As Integer
        Get
            Return propHideEffectDuration
        End Get
    End Property
    
    '    Protected Sub lnkbtnAccessKey_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '        Dim lb As LinkButton = sender
    '        Dim x As String = tbAccessKey.Text.ToLower
    '        If x = "nb" Then
    '            Server.Transfer("NoticeBoard.aspx")
    '        End If
    '        If x = "or" Then
    '            Server.Transfer("on_line_picks.aspx")
    '        End If
    '        If x = "cc" Or x = "cb" Then
    '            Server.Transfer("courier_collection.aspx")
    '        End If
    '        If x = "tt" Then
    '            Server.Transfer("track_and_trace.aspx")
    '        End If
    '        If x = "re" Or x = "rp" Then
    '            Server.Transfer("reports.aspx")
    '        End If
    '        If x = "ab" Then
    '            Server.Transfer("address_book.aspx")
    '        End If
    '        If x = "pm" Then
    '            Server.Transfer("product_manager.aspx")
    '        End If
    '        If x = "um" Then
    '            Server.Transfer("user_manager.aspx")
    '        End If
    '        If x = "mp" Then
    '            Server.Transfer("MyProfile.aspx")
    '        End If
    '        If x = "se" Then
    '            Server.Transfer("SiteEditor.aspx")
    '        End If
    '        If x = "ad" Then
    '            Server.Transfer("Administrator.aspx")
    '        End If
    '        If x = "ah" Then
    '            Server.Transfer("AccountHandler.aspx")
    '        End If
    '        If x = "lo" Then
    '            Server.Transfer("default.aspx")
    '        End If
    '    End Sub

    Protected Sub lnkbtnChangeLoginPassword_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ChangeLoginPassword()
    End Sub

    Protected Sub ChangeLoginPassword()
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String = "UPDATE UserProfile SET MustChangePassword = 1 WHERE [Key] = " & Session("UserKey")
        Dim oCmd As New SqlCommand(sSQL, oConn)
        Try
            oConn.Open()
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
            WebMsgBox.Show("ChangeLoginPassword: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        WebMsgBox.Show("You will be prompted to change your password next time you log into the system")
    End Sub

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

</script>
<style type="text/css">
    .TabView
    {
        color: #000000;
        font-size: 8pt;
        font-family: arial;
    }
    A.TabView:LINK
    {
        text-decoration: none;
        color: #000000;
        font-size: 8pt;
        font-family: arial;
    }
    A.TabView:VISITED
    {
        text-decoration: none;
        color: #000000;
        font-size: 8pt;
        font-family: arial;
    }
    A.TabView:HOVER
    {
        text-decoration: underline;
        color: #4682b4;
        font-size: 8pt;
        font-family: arial;
    }
    
    .TabView_RAMBLERS
    {
        color: #000000;
        font-size: 10pt;
        font-family: arial;
    }
    A.TabView_RAMBLERS:LINK
    {
        text-decoration: none;
        color: #000000;
        font-size: 10pt;
        font-family: arial;
    }
    A.TabView_RAMBLERS:VISITED
    {
        text-decoration: none;
        color: #000000;
        font-size: 10pt;
        font-family: arial;
    }
    A.TabView_RAMBLERS:HOVER
    {
        text-decoration: underline;
        color: #4682b4;
        font-size: 10pt;
        font-family: arial;
    }
</style>
<table style="width: 100%">
    <tr>
        <td style="width: 1%">
        </td>
        <td style="width: 48%">
            <asp:Image runat="server" ID="imgRunningHeader"></asp:Image>
        </td>
        <td style="width: 1%">
            <img border="0" src="../images/blank.gif" width="1" height="1" alt="" />
        </td>
        <td id="tdNewRotator" runat="server" style="width: 50%">
            <asp:Label ID="lblRotator" runat="server" />
        </td>
        <td id="tdOldRotator" runat="server" style="width: 50%">
            <ComponentArt:Rotator ID="Rotator1" runat="server" CssClass="rotatortext" Width="100%" Height="44" ScrollInterval="<%# intScrollInterval %>" RotationType="<%# enumRotationType %>" SmoothScrollSpeed="<%# enumSmoothScrollSpeed %>" PauseOnMouseOver="<%# blnPauseOnMouseOver %>" SlidePause="<%# intSlidePause %>" ScrollDirection="<%# enumScrollDirection %>" ShowEffect="<%# enumShowEffect %>" ShowEffectDuration="<%# intShowEffectDuration %>" HideEffect="<%# enumHideEffect %>" HideEffectDuration="<%# intHideEffectDuration %>" Loop="<%# blnContinuousLoop %>" Visible="<%# blnVisible %>">
                <SlideTemplate>
                    <table cellspacing="1" cellpadding="0" border="0">
                        <tr>
                            <td class="rotatortext">
                                <img src='<%# sVirtualAdrotatorThumbFolder %><%# DataBinder.Eval(Container.DataItem, "thumbnailimage") %> ' height="44" alt="" />&nbsp;
                            </td>
                            <td class="rotatortext">
                                &nbsp;<%# DataBinder.Eval(Container.DataItem, "AdRotatorText") %>
                            </td>
                        </tr>
                    </table>
                </SlideTemplate>
            </ComponentArt:Rotator>
        </td>
    </tr>
</table>
<br />
<VASP:ASPTabView ID="TabView" runat="server" />
<table style="width: 100%; border: none" cellpadding="0" cellspacing="0">
    <tr class="bar_statusline">
        <td style="white-space: nowrap; height: 19px;" align="right">
            <asp:Label ID="lblTopLineMessage" runat="server" ForeColor="Yellow" Font-Bold="True" Font-Size="XX-Small" Font-Names="Verdana"></asp:Label>
            &nbsp;&nbsp;
            <asp:Label ID="lblUserName" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="white"></asp:Label>
            <asp:LinkButton CssClass="TreeView" OnClientClick='window.open("ShowRoles.aspx","ShowRoles","top=100,left=100,width=450,height=250,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=no");' ID="lnkbtnCustomerName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="White" Font-Bold="True" />
            &nbsp;&nbsp;&nbsp;
            <asp:LinkButton runat="server" OnClick="btn_LogOut_click" CausesValidation="false" ForeColor="Blue" ID="btn_LogOut" Font-Underline="false" Font-Size="XX-Small" Font-Names="Verdana" Style="text-decoration: none" Text="[log out]" />
            <asp:LinkButton ID="lnkbtnChangeLoginPassword" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Underline="false" OnClick="lnkbtnChangeLoginPassword_Click" Style="text-decoration: none">[chng pwd]</asp:LinkButton>
            &nbsp;&nbsp;&nbsp;
        </td>
    </tr>
</table>