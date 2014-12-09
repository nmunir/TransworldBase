<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Telerik.Web.UI" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' PRODUCT CREDITS
    ' decrement credits
    ' lblAvailable, imgbtnAvailable
    ' ProductCreditStatus

    
    ' check enforce correct way round in manual refresh
    ' check deleted user groups are ignored
    ' NOTE THAT if order qty reduced because of insufficient available qty, no authorisation for more qty is generated (but a message indicating that the order has been reduced because of insufficient quantity is shown to agent)

    ' need to handle NO consignment being generated (all products requiring authorisation)
    ' we don't need AuthoriserKey in CreateAuthorisationRequest
    
    ' TEST!!!
    ' support multiple destinations
    ' add service level
    ' remove CRLF from Special Instrs & Packing Note fields
    
    ' in some circumstances NEW MESSAGE doesn't flash
    
    ' line 860 - For Each dr As DataRow In gdtBasket.Rows - crash with "Object reference not set to an instance of an object." if refresh after placing order
    
    Const ITEMS_PER_REQUEST As Integer = 30
    Const COUNTRY_UK As Int32 = 222
    Const ACCOUNT_CODE As String = "COURI11111"
    Const LICENSE_KEY As String = "RA61-XZ94-CT55-FH67"

    Const COUNTRY_CODE_CANADA As Int32 = 38
    Const COUNTRY_CODE_USA As Int32 = 223
    Const COUNTRY_CODE_USA_NYC As Int32 = 256
    
    Const CUSTOMER_INTERNAL As Int32 = 566
    
    Const SHOW_ADDRESS_BOOK As Boolean = False
    Const SHOW_SAVEADDRESS As Boolean = False          ' was True
    Const SHOW_PACKINGNOTE As Boolean = False
    
    Const CREDIT_LIMIT_ENFORCE_FALSE As Int32 = 0
    Const CREDIT_LIMIT_ENFORCE_TRUE As Int32 = 1
    Const NO_MAX_GRAB As Int32 = -1
    
    Const ORDER_STATUS_OK As Int32 = 1
    Const ORDER_STATUS_REDUCED_TO_TOTAL_AVAILABLE = 2
    Const ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT = 3
    Const ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT_REQUEST_AUTHORISATION = 4
    Const ORDER_STATUS_REDUCED_TO_MAX_GRAB = 5

    Const LAST_ADDRESS_COOKIE As String = "TW_LastAddress"

    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gdtBasket As DataTable
    Private gdrCnor As DataRow
   
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim rsm As RadScriptManager = ScriptManager.GetCurrent(Me.Page)
        'rsm.EnablePageMethods = True
        If Not IsNumeric(Session("UserKey")) Then
            Response.RedirectLocation = "http:/my.transworld.eu.com/common/session_expired.aspx"
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call InitPage()
            '            If pnNewMessagesMessageShown = 0 Then
            Dim sSQL As String = "SELECT NewMessage FROM MessagingTopics WHERE CreatedOn >= GETDATE() - 30 AND UserKey = " & Session("UserKey") & " AND NewMessage = 1"
            If ExecuteQueryToDataTable(sSQL).Rows.Count > 0 Then
                'WebMsgBox.Show("You have new messages.\n\nPlease view your new messages before placing another order.\n\nThank you.")
                pnlMain.Visible = False
                pnlRedirect.Visible = True
                'Exit Sub
            End If
            'pnNewMessagesMessageShown = 1
            Call CheckForOverdraftRequests()
        End If
        'End If
        'If pnNewMessagesMessageShown = 1 Then
        'Call InitPage()
        'pnNewMessagesMessageShown = 2
        'End If
    End Sub

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sm As New ScriptManager
        sm.ID = "ScriptMgr"
        Try
            PlaceHolderForScriptManager.Controls.Add(sm)
        Catch ex As Exception
        End Try
    End Sub
    
    Protected Sub CheckForOverdraftRequests()
        If GetOverdraftRequests() > 0 Then
            Call SetAuthRequestsVisibility(True)
        Else
            Call SetAuthRequestsVisibility(False)
        End If
    End Sub
    
    Protected Sub InitPage()
        If CInt(Session("CustomerKey")) = CUSTOMER_INTERNAL Then
            Call GetCustomerAccountCodes()
            divInternal.Visible = True
            divMainForm.Visible = False
            ddlCustomer.Focus()
        Else
            pnImpersonateCustomer = CInt(Session("CustomerKey"))
            pnImpersonateBookedByUser = CInt(Session("UserKey"))
            divMainForm.Visible = True
            'Call GetSiteFeatures()
            Call HideDefaultCustRefFields()
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
        Call SetAddressVisibility(False)
        Call BindRotator()
        If Not InitCneeAddress() Then
            Call GetCountries()
            cbSaveAddress.Visible = True
            lblLegendCheckAddress.Visible = False
        Else
            lblLegendCheckAddress.Visible = True
        End If
        If SHOW_ADDRESS_BOOK Then
            lnkbtnShowAddressBook.Visible = True
        Else
            lnkbtnShowAddressBook.Visible = False
        End If
        If SHOW_SAVEADDRESS Then
            cbSaveAddress.Visible = True
        Else
            cbSaveAddress.Visible = False
        End If
        If SHOW_PACKINGNOTE Then
            trPackingNote.Visible = True
        Else
            trPackingNote.Visible = False
        End If
    End Sub
    
    Protected Function GetOverdraftRequests() As Int32
        Dim sSQL As String = "SELECT ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), OrderCreatedDateTime, 106), ' ', '-') + ' ' + SUBSTRING((CONVERT(VARCHAR(8), OrderCreatedDateTime, 108)),1,5) AS varchar(20)),'(never)') 'OrderCreatedDateTime', OrderStatus, [id], ISNULL(ConsignmentKey,0) 'ConsignmentKey' FROM ProductCreditsOrderHoldingQueue WHERE UserProfileKey = " & Session("UserKey") & " AND OrderCreatedDateTime >= GETDATE() - 31 ORDER BY [id] DESC"
        Dim dtOverdraftRequests As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim nRowCount As Int32 = dtOverdraftRequests.Rows.Count
        If nRowCount > 0 Then
            gvOverdraftRequests.DataSource = dtOverdraftRequests
            gvOverdraftRequests.DataBind()
        End If
        GetOverdraftRequests = nRowCount
    End Function
    
    Protected Sub SetAuthRequestsVisibility(ByVal bVisible As Boolean)
        trAuthRequests01.Visible = bVisible
        trAuthRequests02.Visible = bVisible
        trAuthRequests03.Visible = bVisible
    End Sub
    Protected Function GetImage(ByVal DataItem As Object) As String
        
        Dim sVirtualJpegImageUrl As String = ConfigLib.GetConfigItem_Virtual_JPG_URL
        Dim sImageName As String = DataBinder.Eval(DataItem, "ThumbNailImage")
        Return sVirtualJpegImageUrl + sImageName
        
    End Function
    
    <System.Web.Services.WebMethod()>
    Public Shared Function SetCustomTextOnProductDropDown(ByVal nProductKey As Integer) As String
        SetCustomTextOnProductDropDown = String.Empty
        If IsNumeric(nProductKey) Then
            Dim sSQL As String = "SELECT ProductCode 'Product' from LogisticProduct where LogisticProductKey = " & nProductKey
            Dim oDataTable As New DataTable
            Dim oConn As New SqlConnection(ConfigLib.GetConfigItem_ConnectionString)
            Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
            Try
                oConn.Open()
                oAdapter.Fill(oDataTable)
                If Not oDataTable Is Nothing AndAlso oDataTable.Rows.Count > 0 Then
                    SetCustomTextOnProductDropDown = oDataTable.Rows(0)("Product")
                End If
            Catch ex As Exception
                WebMsgBox.Show(ex.Message.ToString)
            Finally
                oConn.Close()
            End Try
        End If
    End Function
    
    Protected Sub BindRotator()
        Dim dt As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_WUQuickOrder_GetMostPopularProducts", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserKey").Value = Session("UserKey")

        oAdapter.Fill(dt)
        
        'GetProductsByCustomer = dt


        'Dim sSQL As String
        'sSQL = "SELECT lp.LogisticProductKey, lp.ThumbNailImage, ProductDescription + ' (' + ProductCode + ')' 'Product' FROM UserProductFavouritesDefaults upfd "
        'sSQL += "INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = upfd.ProductKey WHERE upfd.CustomerKey = " & Session("CustomerKey")
        'Dim oDataTable As New DataTable
        'Dim oConn As New SqlConnection(gsConn)
        'Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        'Try
        '    oConn.Open()
        '    oAdapter.Fill(oDataTable)
        'Catch ex As Exception
        '    WebMsgBox.Show(ex.Message.ToString)
        'Finally
        '    oConn.Close()
        'End Try
        radRotator.DataSource = dt
        radRotator.DataBind()
    End Sub
    
    Protected Sub radRotator_ItemDataBound(ByVal sender As Object, ByVal e As RadRotatorEventArgs) Handles radRotator.ItemDataBound
        
        Dim hidLogisticProductKey As HiddenField = e.Item.FindControl("hidLogisticProductKey")
        Dim imgRotator As System.Web.UI.WebControls.Image = e.Item.FindControl("imgRotator")
        Dim nProductKey As Integer = Convert.ToInt32(hidLogisticProductKey.Value)
        imgRotator.Attributes.Add("onclick", "ImageClick('" & nProductKey & "')")
        
    End Sub
    
    Protected Function InitCneeAddress() As Boolean
        InitCneeAddress = False
        Dim sSQL As String
        Dim sUserID As String = ExecuteQueryToDataTable("SELECT UserID FROM UserProfile WHERE [key] = " & Session("UserKey")).Rows(0).Item(0)
        sSQL = "SELECT * FROM ClientData_WU_Agents WHERE Termid = '" & sUserID & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            InitCneeAddress = True
            cbSaveAddress.Visible = False
            trCountry.Visible = False
            trCneeName.Visible = False
            lblLegendContactName.ForeColor = Drawing.Color.Red
            
            Dim dr As DataRow = dt.Rows(0)
            trCneeNameReadOnly.Visible = True
            lblCneeNameReadOnly.Text = dr("AgentName")
            trCneeAddr1ReadOnly.Visible = True
            lblCneeAddr1ReadOnly.Text = dr("Address1")
            lblCneeAddr2ReadOnly.Text = dr("Address2").ToString.Trim
            If lblCneeAddr2ReadOnly.Text <> String.Empty Then
                trCneeAddr2ReadOnly.Visible = True
            End If
            lblCneeAddr3ReadOnly.Text = dr("Address3").ToString.Trim
            If lblCneeAddr3ReadOnly.Text <> String.Empty Then
                trCneeAddr3ReadOnly.Visible = True
            End If
            trCneeTownCityReadOnly.Visible = True
            lblCneeTownCityReadOnly.Text = dr("City")
            lblCneeStateReadOnly.Text = dr("State").ToString.Trim
            If lblCneeStateReadOnly.Text <> String.Empty Then
                trCneeStateReadOnly.Visible = True
            End If
            trCneePostcodeReadOnly.Visible = True
            lblCneePostcodeReadOnly.Text = dr("Postcode")
            rfvCneeCtcName.Enabled = True
        End If
    End Function

    Protected Sub CreateOrRetrieveCookie()
        If Request.Cookies(LAST_ADDRESS_COOKIE) Is Nothing Then
            Call CreateLastAddressCookie()
        Else
            tbCneeCtcName.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("ContactName") & String.Empty
            tbCneeName.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Name") & String.Empty
            tbCneeAddr1.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Addr1") & String.Empty
            tbCneeAddr2.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Addr2") & String.Empty
            tbCneeAddr3.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Addr3") & String.Empty
            tbCneeTown.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Town") & String.Empty
            tbCneeState.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("State") & String.Empty
            tbCneePostCode.Text = Request.Cookies(LAST_ADDRESS_COOKIE)("Postcode") & String.Empty
            Dim sCneeCountryCode As String = Request.Cookies(LAST_ADDRESS_COOKIE)("CountryCode") & String.Empty
            If IsNumeric(sCneeCountryCode) Then
                For i As Int32 = 1 To ddlCountry.Items.Count - 1
                    If ddlCountry.Items(i).Value = sCneeCountryCode Then
                        ddlCountry.SelectedIndex = i
                        Call SetAddressVisibility(True)
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
   
    Protected Sub CreateLastAddressCookie()
        Dim c As HttpCookie = New HttpCookie(LAST_ADDRESS_COOKIE)
        c.Values.Add("ContactName", String.Empty)
        c.Values.Add("Name", String.Empty)
        c.Values.Add("Addr1", String.Empty)
        c.Values.Add("Addr2", String.Empty)
        c.Values.Add("Addr3", String.Empty)
        c.Values.Add("Town", String.Empty)
        c.Values.Add("State", String.Empty)
        c.Values.Add("Postcode", String.Empty)
        c.Values.Add("CountryCode", String.Empty)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub

    Protected Sub HideDefaultCustRefFields()
        trCustRef1.Visible = False
        trCustRef2.Visible = False
        trCustRef3.Visible = False
        trCustRef4.Visible = False
    End Sub
    
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
            trCustRef4.Visible = False
        End If
    End Sub

    Protected Sub SetAddressVisibility(ByVal bVisible As Boolean)
        trPostCode.Visible = bVisible
        trCneeAddr1.Visible = bVisible
        trCneeAddr2.Visible = bVisible
        trCneeAddr3.Visible = bVisible
        trTownCity.Visible = bVisible
        trCneeState.Visible = bVisible
        trPostCode.Visible = bVisible
    End Sub
                                      
    Protected Sub btnAddToOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        If rcbProduct.SelectedValue = String.Empty Then
            Exit Sub
        End If
        
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
        sSQL = "SELECT ISNULL(MinGrab, 0) FROM ClientData_WU_MinGrabs WHERE LogisticProductKey = " & rcbProduct.SelectedValue & " AND UserGroup = (SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & pnImpersonateBookedByUser & ")"
        Dim dtMinGrab As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim nMinGrab As Int32 = 0
        If dtMinGrab.Rows.Count > 0 Then
            nMinGrab = dtMinGrab.Rows(0).Item(0)
        End If
        If nMinGrab > CInt(rntbQty.Text) Then
            WebMsgBox.Show("This product has a minimum order quantity of " & nMinGrab.ToString & ".\n\nPlease amend your order.")
            rntbQty.Focus()
            Exit Sub
        End If

        Call CreateBasketIfNull()
        gdtBasket = Session("BO_BasketData")
        Dim gdvBasket As New DataView(gdtBasket)
        gdvBasket.RowFilter = "LogisticProductKey='" & rcbProduct.SelectedValue & "'"
        If gdvBasket.Count > 0 Then
            WebMsgBox.Show("This product is already in your basket.\n\nTo change the quantity, remove it from the basket and re-select the product with the quantity required.")
        Else
            Dim dr As DataRow = gdtBasket.NewRow()
            dr("LogisticProductKey") = rcbProduct.SelectedValue
            dr("Product") = rcbProduct.Text
            dr("TotalAvailable") = GetTotalAvailableQty(rcbProduct.SelectedValue)
            dr("MaxGrab") = GetMaxGrab(rcbProduct.SelectedValue)
            dr("QtyRequested") = rntbQty.Text
            dr("Message") = String.Empty

            sSQL = "SELECT ISNULL(ProductCredits, 0) FROM LogisticProduct WHERE LogisticProductKey = " & rcbProduct.SelectedValue
            Dim bUsesProductCredits As Boolean = CBool(ExecuteQueryToDataTable(sSQL).Rows(0).Item(0))
            Dim dtProductCredits As DataTable = Nothing
            Dim drProductCredits As DataRow = Nothing
            If bUsesProductCredits Then
                dtProductCredits = ExecuteQueryToDataTable("SELECT TOP 1 * FROM ProductCredits WHERE LogisticProductKey = " & rcbProduct.SelectedValue & " AND UserKey = " & Session("UserKey") & " AND GETDATE() > CreditStartDateTime AND GETDATE() < CreditEndDateTime ORDER BY [id] DESC")
                If dtProductCredits.Rows.Count = 0 Then
                    bUsesProductCredits = False
                Else
                    drProductCredits = dtProductCredits.Rows(0)
                End If
            End If
            If bUsesProductCredits Then
                dr("UsesProductCredits") = True
                dr("CreditRecordID") = drProductCredits("id")
                'Dim dtProductCredits As DataTable = ExecuteQueryToDataTable("SELECT TOP 1 * FROM ProductCredits WHERE LogisticProductKey = " & rcbProduct.SelectedValue & " AND UserKey = " & Session("UserKey") & " ORDER BY [id] DESC")
                'If dtProductCredits.Rows.Count > 0 Then
                dr("RemainingCredit") = drProductCredits("RemainingCredit")
                dr("EnforceCreditLimit") = drProductCredits("EnforceCreditLimit")
                Dim dtCreditStartDateTime As DateTime = drProductCredits("CreditStartDateTime")
                Dim dtCreditEndDateTime As DateTime = drProductCredits("CreditEndDateTime")
                If dtCreditStartDateTime > DateTime.Now Then
                    dr("RefreshDate") = dtCreditStartDateTime.ToString("dd-MMM-yyyy hh:mm")
                ElseIf dtCreditEndDateTime > DateTime.Now Then
                    dr("RefreshDate") = dtCreditEndDateTime.ToString("dd-MMM-yyyy hh:mm")
                Else
                    dr("RefreshDate") = "(no refresh date found)"
                End If
                'Else
                '    dr("RemainingCredit") = 0
                '    dr("EnforceCreditLimit") = 0
                '    dr("RefreshDate") = "(no refresh date found)"
                'End If
            Else
                dr("UsesProductCredits") = False
            End If
            
            Dim nAvailableQuantity As Int32
            
            If bUsesProductCredits Then
                ' USING CREDITS
                nAvailableQuantity = Math.Min(dr("TotalAvailable"), dr("RemainingCredit"))
                If rntbQty.Text <= nAvailableQuantity Then
                    ' QUANTITY REQUESTED <= QUANTITY AVAILABLE
                    dr("QtyGranted") = rntbQty.Text
                    dr("OrderStatus") = ORDER_STATUS_OK
                    gdtBasket.Rows.Add(dr)
                Else
                    ' QUANTITY REQUESTED > QUANTITY AVAILABLE
                    If dr("EnforceCreditLimit") = CREDIT_LIMIT_ENFORCE_TRUE Then
                        ' ENFORCE
                        If dr("TotalAvailable") >= nAvailableQuantity Then
                            ' exceeded credit
                            dr("QtyGranted") = nAvailableQuantity
                            dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT
                            dr("Message") = "You can order a maximum of " & nAvailableQuantity.ToString & " in the current order period, which lasts until " & CDate(dr("RefreshDate")).ToString("dd-MMM-yyyy") & ", after which may be able to order more."
                            gdtBasket.Rows.Add(dr)
                        Else
                            ' qty not available
                            If dr("TotalAvailable") > 0 Then
                                dr("QtyGranted") = nAvailableQuantity
                                dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_TOTAL_AVAILABLE
                                dr("Message") = "The quantity you requested has been reduced to the quantity currently available.<br /><br />Please order again later if you require more."
                                gdtBasket.Rows.Add(dr)
                            Else
                                ' none available
                                WebMsgBox.Show("None of the product is currently available.")
                            End If
                        End If
                    Else
                        ' OVERDRAFT
                        If dr("TotalAvailable") >= nAvailableQuantity Then
                            ' exceeded credit
                            dr("QtyGranted") = nAvailableQuantity
                            dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT_REQUEST_AUTHORISATION
                            dr("Message") = "The quantity you requested will be reduced to " & dr("QtyGranted") & ".<br /><br />This is your maximum remaining allowance for the current order period, which lasts until " & CDate(dr("RefreshDate")).ToString("dd-MMM-yyyy") & ".<br /><br />If you proceed with the order, a request for the excess amount of " & (dr("QtyRequested") - dr("QtyGranted")) & " will be submitted for authorisation.<br /><br />If approved, this will be sent to you separately."
                            gdtBasket.Rows.Add(dr)
                        Else
                            ' qty not available
                            If dr("TotalAvailable") > 0 Then
                                dr("QtyGranted") = dr("TotalAvailable")
                                dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_TOTAL_AVAILABLE
                                dr("Message") = "The quantity you requested has been reduced to the quantity currently available.<br /><br />Please order again later if you require more."
                                gdtBasket.Rows.Add(dr)
                            Else
                                ' none available
                                WebMsgBox.Show("None of this product is currently available.")
                            End If
                        End If
                    End If
                End If
            Else
                ' NOT USING CREDITS
                If dr("MaxGrab") = NO_MAX_GRAB Then
                    nAvailableQuantity = dr("TotalAvailable")
                Else
                    nAvailableQuantity = Math.Min(dr("TotalAvailable"), dr("MaxGrab"))
                End If
                If rntbQty.Text <= nAvailableQuantity Then
                    ' QUANTITY REQUESTED <= QUANTITY AVAILABLE
                    dr("QtyGranted") = rntbQty.Text
                    dr("OrderStatus") = ORDER_STATUS_OK
                    gdtBasket.Rows.Add(dr)
                Else
                    ' QUANTITY REQUESTED > QUANTITY AVAILABLE
                    If dr("TotalAvailable") >= nAvailableQuantity Then
                        dr("QtyGranted") = nAvailableQuantity
                        dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_MAX_GRAB
                        dr("Message") = "The quantity you requested has been reduced to the maximum allowed order quantity.<br /><br />Please order again later if you require more."
                        gdtBasket.Rows.Add(dr)
                    Else
                        ' qty not available
                        If dr("TotalAvailable") > 0 Then
                            dr("QtyGranted") = dr("TotalAvailable")
                            dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_TOTAL_AVAILABLE
                            dr("Message") = "The quantity you requested has been reduced to the quantity currently available.<br /><br />Please order again later if you require more."
                            gdtBasket.Rows.Add(dr)
                        Else
                            ' none available
                            WebMsgBox.Show("None of this product is currently available.")
                        End If
                    End If
                End If
            End If
        End If
        
        Session("BO_BasketData") = gdtBasket
        gvBasket.DataSource = gdtBasket
        gvBasket.DataBind()
        Call SetPlaceOrderButtonVisibility()
        rcbProduct.SelectedIndex = 0
        rntbQty.Text = "1"
        
        ddlCustomer.Enabled = False
        btnAddToOrder.Enabled = False
        rntbQty.Enabled = False
        rcbProduct.Focus()
        Call ClearOrderConfirmation()
        
        rcbProduct.Text = String.Empty
    End Sub

    Protected Sub AdjustUserCredits()
        Dim sSQL As String
        gdtBasket = Session("BO_BasketData")
        For Each dr As DataRow In gdtBasket.Rows
            If CBool(dr("UsesProductCredits")) Then
                ' dr("QtyGranted") 
                ' dr("LogisticProductKey")
                sSQL = "UPDATE ProductCredits SET RemainingCredit = RemainingCredit - " & dr("QtyGranted") & " WHERE [id] = " & dr("CreditRecordID")
                Call ExecuteQueryToDataTable(sSQL)
            End If
        Next
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
            gdtBasket.Columns.Add(New DataColumn("TotalAvailable", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("MaxGrab", GetType(Int32)))
            'gdtBasket.Columns.Add(New DataColumn("Available", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("UsesProductCredits", GetType(Boolean)))
            gdtBasket.Columns.Add(New DataColumn("CreditRecordID", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("StartCredit", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("RemainingCredit", GetType(Int32)))            ' amount of product that can be ordered, same as available unless lowered by product credits
            gdtBasket.Columns.Add(New DataColumn("EnforceCreditLimit", GetType(Int32)))  ' 0 - ENFORCE, 1 = more if authorised
            gdtBasket.Columns.Add(New DataColumn("RefreshDate", GetType(String)))         ' blank if no refresh date, otherwise latest date on which product credit is refreshed
            gdtBasket.Columns.Add(New DataColumn("QtyRequested", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("QtyGranted", GetType(Int32)))
            gdtBasket.Columns.Add(New DataColumn("RequiresAuthorisation", GetType(Boolean)))
            gdtBasket.Columns.Add(New DataColumn("OrderStatus", GetType(Int32)))  ' 0 
            gdtBasket.Columns.Add(New DataColumn("Message", GetType(String)))
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
            'Call PopulateProductDropdown(ddlCustomer.SelectedValue)            ' XXXX
            pnImpersonateCustomer = ddlCustomer.SelectedValue
            Call GetBookedByUsers()
            ddlBookedBy.Enabled = True
            ddlBookedBy.Focus()
        Else
            ddlBookedBy.Enabled = False
        End If
        Call ClearOrderConfirmation()
    End Sub

    Protected Function GetMaxGrab(ByVal sLogisticProductKey As String) As Int32
        GetMaxGrab = -1
        If Session("UserType").ToString.ToLower <> "superuser" Then
            Dim dtUserProductProfile As DataTable = ExecuteQueryToDataTable("SELECT ApplyMaxGrab, MaxGrabQty FROM UserProductProfile WHERE UserKey = " & Session("UserKey") & " AND ProductKey = " & sLogisticProductKey)
            If dtUserProductProfile.Rows.Count > 0 Then
                Dim dr As DataRow = dtUserProductProfile.Rows(0)
                If dr("ApplyMaxGrab") <> 0 Then
                    GetMaxGrab = dr("MaxGrabQty")
                End If
            End If
        End If
    End Function

    Protected Function GetTotalAvailableQty(ByVal sLogisticProductKey As String) As Int32
        GetTotalAvailableQty = ExecuteQueryToDataTable("SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END").Rows(0).Item(0)
    End Function

    Protected Function GetMinOfMaxGrabAndTotalAvailable(ByVal sLogisticProductKey As String) As Int32
        GetMinOfMaxGrabAndTotalAvailable = 0
        Dim nMaxGrab As Int32 = GetMaxGrab(sLogisticProductKey)
        Dim nTotalAvailable As Int32 = GetTotalAvailableQty(sLogisticProductKey)
        If nMaxGrab < 0 Then
            GetMinOfMaxGrabAndTotalAvailable = nTotalAvailable
        Else
            If nMaxGrab < nTotalAvailable Then
                GetMinOfMaxGrabAndTotalAvailable = nMaxGrab
            Else
                GetMinOfMaxGrabAndTotalAvailable = nTotalAvailable
            End If
        End If
    End Function
    
    'Protected Function GetAvailableQty(ByVal sLogisticProductKey As String) As Int32
    '    'Dim sSQL As String = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END"
    '    Dim nApplyMaxGrab As Int32 = 0
    '    Dim nMaxGrabQty As Int32 = 0
    '    Dim sSQL As String
    '    If Session("UserType").ToString.ToLower <> "superuser" Then
    '        sSQL = "SELECT ApplyMaxGrab, MaxGrabQty FROM UserProductProfile WHERE UserKey = " & Session("UserKey") & " AND ProductKey = " & sLogisticProductKey
    '        Dim dtUserProductProfile As DataTable = ExecuteQueryToDataTable(sSQL)
    '        If dtUserProductProfile.Rows.Count > 0 Then
    '            Dim dr As DataRow = dtUserProductProfile.Rows(0)
    '            nApplyMaxGrab = dr("ApplyMaxGrab")
    '            nMaxGrabQty = dr("MaxGrabQty")
    '        End If
    '    End If
    '    sSQL = "SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END"
    '    Dim nTotalQtyAvailable = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    '    Try
    '        If nApplyMaxGrab <> 0 Then
    '            If nMaxGrabQty < nTotalQtyAvailable Then
    '                GetAvailableQty = nMaxGrabQty
    '            Else
    '                GetAvailableQty = nTotalQtyAvailable
    '            End If
    '        Else
    '            GetAvailableQty = nTotalQtyAvailable
    '        End If
    '    Catch
    '        GetAvailableQty = 0
    '    End Try
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
        oAdapter.SelectCommand.Parameters("@FavouriteProducts").Value = IIf(rbFavouriteProducts.Checked, 1, 0)

        oAdapter.Fill(dt)
        
        GetProductsByCustomer = dt
        Exit Function
        
        Dim sSQL As String
        'sSQL = "SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' ' + ProductDescription 'Product', LogisticProductKey, ThumbnailImage FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation lpl ON lp.LogisticProductKey = lpl.LogisticProductKey INNER JOIN UserProductProfile As upp ON lp.LogisticProductKey = upp.ProductKey WHERE lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' AND lp.CustomerKey = " & sCustomerKey
        'sSQL = "SELECT lp.ProductCode + ' ' + ISNULL(lp.ProductDate,'') + ' ' + lp.ProductDescription 'Product', lp.LogisticProductKey, lp.ThumbnailImage, Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation lpl ON lp.LogisticProductKey = lpl.LogisticProductKey INNER JOIN UserProductProfile As upp ON lp.LogisticProductKey = upp.ProductKey WHERE lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' AND upp.AbleToPick = 1 AND upp.UserKey = " & Session("UserKey")

        ' Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END 

        'sSQL = "SELECT lp.ProductCode + ' ' + ISNULL(lp.ProductDate,'') + ' ' + lp.ProductDescription 'Product', lp.LogisticProductKey, lp.ThumbnailImage FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation lpl ON lp.LogisticProductKey = lpl.LogisticProductKey INNER JOIN UserProductProfile As upp ON lp.LogisticProductKey = upp.ProductKey WHERE lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' AND upp.AbleToPick = 1 AND upp.UserKey = " & Session("UserKey")
        sSQL = "SELECT lp.ProductCode + ' ' + ISNULL(lp.ProductDate,'') + ' ' + lp.ProductDescription 'Product', lp.LogisticProductKey, lp.ThumbnailImage, Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END FROM LogisticProduct lp LEFT OUTER JOIN LogisticProductLocation lpl ON lp.LogisticProductKey = lpl.LogisticProductKey INNER JOIN UserProductProfile As upp ON lp.LogisticProductKey = upp.ProductKey WHERE lp.ArchiveFlag = 'N' AND lp.DeletedFlag = 'N' AND upp.AbleToPick = 1 AND upp.UserKey = " & Session("UserKey")
        If sFilter <> String.Empty Then
            sFilter = sFilter.Replace("'", "''")
            Dim sFilterPurged As String = sFilter
            If sFilterPurged.Contains("max order") Then
                Dim nStart As Int32 = sFilterPurged.IndexOf(" (max order")
                sFilterPurged = sFilterPurged.Substring(0, nStart)
            End If
            sSQL += " AND (lp.ProductCode LIKE '%" & sFilterPurged & "%' OR lp.ProductDescription LIKE '%" & sFilterPurged & "%')"
        Else
            sFilter = "_"
        End If
        ' tbSearch.Text = tbSearch.Text.Trim
        'If psSearchString <> String.Empty Then
        '    sSQL += " AND (ProductCode LIKE '%" & psSearchString & "%' OR ProductDescription LIKE '%" & psSearchString & "%')"
        'End If
        If rbFavouriteProducts.Checked Then
            sSQL += " AND (lp.LogisticProductKey IN (SELECT ProductKey FROM UserProductFavouritesDefaults WHERE CustomerKey = " & pnImpersonateCustomer & ")"
            sSQL += " OR lp.LogisticProductKey IN (SELECT ProductKey FROM UserProductFavourites WHERE UserKey = " & pnImpersonateBookedByUser & "))"
        End If
        sSQL += " ORDER BY lp.ProductCode"
        dt = ExecuteQueryToDataTable(sSQL)
        
        For Each dr As DataRow In dt.Rows
            dr("Product") = dr("Product") & " (max order: " & dr("Quantity") & ")"
        Next
        GetProductsByCustomer = dt
    End Function
    
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
        If Not trCneeAddr1ReadOnly.Visible Then
            Call NormaliseCneeAddress()
        End If
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

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        'param13.Value = psCnorCompany
        param13.Value = gdrCnor("CustomerName")
        oCmdAddBooking.Parameters.Add(param13)
        
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = gdrCnor("CustomerAddr1")
        oCmdAddBooking.Parameters.Add(param14)
        
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = gdrCnor("CustomerAddr2")
        oCmdAddBooking.Parameters.Add(param15)
        
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = gdrCnor("CustomerAddr3")
        oCmdAddBooking.Parameters.Add(param16)
        
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = gdrCnor("CustomerTown")
        oCmdAddBooking.Parameters.Add(param17)
        
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = gdrCnor("CustomerCounty")
        oCmdAddBooking.Parameters.Add(param18)
        
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = gdrCnor("CustomerPostCode")
        oCmdAddBooking.Parameters.Add(param19)
        
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = gdrCnor("CustomerCountryKey")
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
        If trCneeNameReadOnly.Visible Then
            param25.Value = lblCneeNameReadOnly.Text
        Else
            param25.Value = tbCneeName.Text
        End If
        oCmdAddBooking.Parameters.Add(param25)

        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            param26.Value = lblCneeAddr1ReadOnly.Text
        Else
            param26.Value = tbCneeAddr1.Text
        End If
        oCmdAddBooking.Parameters.Add(param26)
        
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeAddr2ReadOnly.Visible Then
                param27.Value = lblCneeAddr2ReadOnly.Text
            Else
                param27.Value = String.Empty
            End If
        Else
            param27.Value = tbCneeAddr2.Text
        End If
        oCmdAddBooking.Parameters.Add(param27)

        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeAddr3ReadOnly.Visible Then
                param28.Value = lblCneeAddr3ReadOnly.Text
            Else
                param28.Value = String.Empty
            End If
        Else
            param28.Value = tbCneeAddr3.Text
        End If
        oCmdAddBooking.Parameters.Add(param28)
        
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        If trCneeTownCityReadOnly.Visible Then
            param29.Value = lblCneeTownCityReadOnly.Text
        Else
            param29.Value = tbCneeTown.Text
        End If
        oCmdAddBooking.Parameters.Add(param29)
        
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeStateReadOnly.Visible Then
                param30.Value = lblCneeStateReadOnly.Text
            Else
                param30.Value = String.Empty
            End If
        Else
            param30.Value = tbCneeState.Text
        End If
        oCmdAddBooking.Parameters.Add(param30)
        
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        If trCneePostcodeReadOnly.Visible Then
            param31.Value = lblCneePostcodeReadOnly.Text
        Else
            param31.Value = tbCneePostCode.Text
        End If
        oCmdAddBooking.Parameters.Add(param31)

        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        If trCneeAddr1ReadOnly.Visible Then
            param32.Value = COUNTRY_UK
        Else
            param32.Value = ddlCountry.SelectedValue
        End If
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
                'gdtBasket = Session("BO_BasketData")
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
                    'param56.Value = dr("QtyRequested")
                    param56.Value = dr("QtyGranted")
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
                'gdtBasket = Nothing
                'Session("BO_BasketData") = gdtBasket
                'Call CreateBasketIfNull()
                'btnPlaceOrder.
            Else
                BookingFailed = True
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                nSubmitConsignment = lConsignmentKey
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            oTrans.Rollback("AddBooking")
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

    Protected Sub btnPlaceOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        mpe.PopupControlID = "divConfirmOrder"
        mpe.Show()
    End Sub
   
    Protected Function CheckSufficientQuantityAvailable() As String
        CheckSufficientQuantityAvailable = String.Empty
        gdtBasket = Session("BO_BasketData")
        For Each dr As DataRow In gdtBasket.Rows
            Dim nQtyAvailable As Int32 = GetQuantityAvailable(dr("LogisticProductKey"))
            If nQtyAvailable < dr("QtyRequested") Then
                CheckSufficientQuantityAvailable = "Product """ & dr("Product") & """ has a requested quantity of " & dr("QtyRequested") & " but only " & nQtyAvailable & " is/are available."
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
        Try
            ddlCountry.SelectedIndex = 0
        Catch ex As Exception
        End Try
        tbPackingNote.Text = String.Empty
        tbSpecialInstructions.Text = String.Empty
        rcbProduct.Text = String.Empty
        gdtBasket = Nothing
        Session("BO_BasketData") = gdtBasket
    End Sub
   
    Protected Sub lnkbtnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ' remove item
        Dim lb As LinkButton = sender
        Dim sLogisticProductKey As String = lb.CommandArgument
        Call RemoveItemFromBasket(sLogisticProductKey)
    End Sub
        
    Protected Sub RemoveItemFromBasket(ByVal sLogisticProductKey As String)
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
    
    Property pnNewMessagesMessageShown() As Int32
        Get
            Dim o As Object = ViewState("QO_NewMessagesMessageShown")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("QO_NewMessagesMessageShown") = Value
        End Set
    End Property

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
    
    'Property psSearchString() As String
    '    Get
    '        Dim o As Object = ViewState("QO_SearchString")
    '        If o Is Nothing Then
    '            Return ""
    '        End If
    '        Return CStr(o)
    '    End Get
    '    Set(ByVal Value As String)
    '        ViewState("QO_SearchString") = Value
    '    End Set
    'End Property
  
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
    
    Protected Sub ProcessAuthorisationRequests(ByVal nAssociatedConsignmentNo As Int32)
        Dim nAuthorisationRequestCount As Int32 = 0
        Dim sbAuthorisationEmail As New StringBuilder
        Dim dictAuthorisationRequestItems As New Dictionary(Of Int32, Int32)
        
        For Each dr As DataRow In gdtBasket.Rows
            If dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT_REQUEST_AUTHORISATION Then
                nAuthorisationRequestCount += 1
                dictAuthorisationRequestItems.Add(dr("LogisticProductKey"), CInt(CInt(dr("QtyRequested")) - CInt(dr("QtyGranted"))))
            End If
        Next
        
        If nAuthorisationRequestCount > 0 Then
            Dim guidAuthorisationGuid As Guid = CreateAuthorisationRequest(dictAuthorisationRequestItems, nAssociatedConsignmentNo)
            sbAuthorisationEmail.Append("User ")
            sbAuthorisationEmail.Append(GetUserDetailsFromUserKey(Session("UserKey")))
            sbAuthorisationEmail.Append(Environment.NewLine)
            sbAuthorisationEmail.Append(Environment.NewLine)

            sbAuthorisationEmail.Append("Order Placed: ")
            sbAuthorisationEmail.Append(Date.Now.ToString("dd-MMM-yyyy hh:mm:ss"))
            sbAuthorisationEmail.Append(" Ref: ")
            sbAuthorisationEmail.Append(guidAuthorisationGuid.ToString)
            sbAuthorisationEmail.Append(Environment.NewLine)

            If nAssociatedConsignmentNo > 0 Then
                sbAuthorisationEmail.Append("Associated consignment number: ")
                sbAuthorisationEmail.Append(nAssociatedConsignmentNo.ToString)
            Else
                sbAuthorisationEmail.Append("No associated consignment number.")
            End If
            sbAuthorisationEmail.Append(Environment.NewLine)
            sbAuthorisationEmail.Append(Environment.NewLine)

            sbAuthorisationEmail.Append("The following ")
            sbAuthorisationEmail.Append(nAuthorisationRequestCount.ToString)
            sbAuthorisationEmail.Append(" product(s) require authorisation:")
            sbAuthorisationEmail.Append(Environment.NewLine)
            sbAuthorisationEmail.Append(Environment.NewLine)
            
            For Each kv As KeyValuePair(Of Int32, Int32) In dictAuthorisationRequestItems
                sbAuthorisationEmail.Append("Product: ")
                Dim sProductDetails() As String = GetProductDetailsFromProductKey(kv.Key)
                sbAuthorisationEmail.Append(sProductDetails(0) & " " & sProductDetails(1) & " " & sProductDetails(2))
                sbAuthorisationEmail.Append("Qty: ")
                sbAuthorisationEmail.Append(kv.Value)
                sbAuthorisationEmail.Append(Environment.NewLine)
            Next
            
            Dim sAuthorisationEmail As String = sbAuthorisationEmail.ToString
            Dim dtAuthorisers As DataTable = ExecuteQueryToDataTable("SELECT up.EmailAddr FROM UserProfile up INNER JOIN ProductCreditAuthorisers pca ON up.[key] = pca.UserKey WHERE up.CustomerKey = " & Session("CustomerKey") & " AND pca.AuthoriserType <= 2")
            If dtAuthorisers.Rows.Count > 0 Then
                For Each dr As DataRow In dtAuthorisers.Rows
                    Call SendMail("PROD_CREDIT_AUTH_REQ", dr(0), "Product Credit Overdraft - Authorisation Request", sAuthorisationEmail, sAuthorisationEmail.Replace(Environment.NewLine, "<br />" & Environment.NewLine))
                Next
            End If
        End If
    End Sub

    Protected Function GetUserDetailsFromUserKey(ByVal nUserKey As Int32) As String
        GetUserDetailsFromUserKey = ExecuteQueryToDataTable("SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' FROM UserProfile WHERE [key] = " & nUserKey).Rows(0).Item(0)
    End Function

    Protected Function GetProductDetailsFromProductKey(ByVal nProductKey As Int32) As String()
        Dim sProductDetails() As String = {String.Empty, String.Empty, String.Empty}
        Dim sSQL As String = "SELECT ProductCode, ISNULL(ProductDate,'') 'ProductDate', ISNULL(ProductDescription,'') 'ProductDescription' FROM LogisticProduct WHERE LogisticProductKey = " & nProductKey
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        sProductDetails(0) = dr("ProductCode")
        sProductDetails(1) = dr("ProductDate")
        sProductDetails(2) = dr("ProductDescription")
        GetProductDetailsFromProductKey = sProductDetails
    End Function
    
    Protected Function CreateAuthorisationRequest(ByVal dicOrderItems As Dictionary(Of Integer, Integer), ByVal nOriginalConsignmentKey As Int32) As Guid
        Dim nHoldingQueueKey As Int32
        Dim guidAuthorisationGUID As Guid
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_ProductCredits_AuthorisationRequest", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure

        guidAuthorisationGUID = Guid.NewGuid
        Dim paramAuthorisationGUID As SqlParameter = New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20)
        paramAuthorisationGUID.Value = guidAuthorisationGUID.ToString
        oCmdAddBooking.Parameters.Add(paramAuthorisationGUID)

        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CLng(Session("UserKey"))
        oCmdAddBooking.Parameters.Add(param1)

        Dim paramAuthoriserKey As SqlParameter = New SqlParameter("@AuthoriserKey", SqlDbType.Int, 4)
        paramAuthoriserKey.Value = 0
        oCmdAddBooking.Parameters.Add(paramAuthoriserKey)

        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = pnImpersonateCustomer
        oCmdAddBooking.Parameters.Add(param2)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = tbCustRef1.Text
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = tbCustRef2.Text
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = tbCustRef3.Text
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = tbCustRef4.Text
        oCmdAddBooking.Parameters.Add(param6)
        
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = nOriginalConsignmentKey.ToString
        oCmdAddBooking.Parameters.Add(param6a)

        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = tbSpecialInstructions.Text.Replace(Environment.NewLine, " ").Trim
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = tbPackingNote.Text.Replace(Environment.NewLine, " ").Trim
        oCmdAddBooking.Parameters.Add(param8)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = gdrCnor("CustomerName")
        oCmdAddBooking.Parameters.Add(param13)

        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = gdrCnor("CustomerAddr1")
        oCmdAddBooking.Parameters.Add(param14)

        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = gdrCnor("CustomerAddr2")
        oCmdAddBooking.Parameters.Add(param15)

        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = gdrCnor("CustomerAddr3")
        oCmdAddBooking.Parameters.Add(param16)

        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = gdrCnor("CustomerTown")
        oCmdAddBooking.Parameters.Add(param17)

        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = gdrCnor("CustomerCounty")
        oCmdAddBooking.Parameters.Add(param18)

        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = gdrCnor("CustomerPostCode")
        oCmdAddBooking.Parameters.Add(param19)

        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = gdrCnor("CustomerCountryKey")
        oCmdAddBooking.Parameters.Add(param20)

        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param21)

        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param22)

        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param23)

        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        If trCneeNameReadOnly.Visible Then
            param25.Value = lblCneeNameReadOnly.Text
        Else
            param25.Value = tbCneeName.Text
        End If
        oCmdAddBooking.Parameters.Add(param25)

        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            param26.Value = lblCneeAddr1ReadOnly.Text
        Else
            param26.Value = tbCneeAddr1.Text
        End If
        oCmdAddBooking.Parameters.Add(param26)

        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeAddr2ReadOnly.Visible Then
                param27.Value = lblCneeAddr2ReadOnly.Text
            Else
                param27.Value = String.Empty
            End If
        Else
            param27.Value = tbCneeAddr2.Text
        End If
        oCmdAddBooking.Parameters.Add(param27)

        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeAddr3ReadOnly.Visible Then
                param28.Value = lblCneeAddr3ReadOnly.Text
            Else
                param28.Value = String.Empty
            End If
        Else
            param28.Value = tbCneeAddr3.Text
        End If
        oCmdAddBooking.Parameters.Add(param28)

        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        If trCneeTownCityReadOnly.Visible Then
            param29.Value = lblCneeTownCityReadOnly.Text
        Else
            param29.Value = tbCneeTown.Text
        End If
        oCmdAddBooking.Parameters.Add(param29)

        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        If trCneeAddr1ReadOnly.Visible Then
            If trCneeStateReadOnly.Visible Then
                param30.Value = lblCneeStateReadOnly.Text
            Else
                param30.Value = String.Empty
            End If
        Else
            param30.Value = tbCneeState.Text
        End If
        oCmdAddBooking.Parameters.Add(param30)

        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        If trCneePostcodeReadOnly.Visible Then
            param31.Value = lblCneePostcodeReadOnly.Text
        Else
            param31.Value = tbCneePostCode.Text
        End If
        oCmdAddBooking.Parameters.Add(param31)

        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        If trCneeAddr1ReadOnly.Visible Then
            param32.Value = COUNTRY_UK
        Else
            param32.Value = ddlCountry.SelectedValue
        End If
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

        Dim param36 As SqlParameter = New SqlParameter("@MsgToAuthoriser", SqlDbType.NVarChar, 1000)
        param36.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param36)

        Dim param37 As SqlParameter = New SqlParameter("@HoldingQueueKey", SqlDbType.Int, 4)
        param37.Direction = ParameterDirection.Output
        oCmdAddBooking.Parameters.Add(param37)

        Try
            oConn.Open()
            oCmdAddBooking.ExecuteNonQuery()
            nHoldingQueueKey = CInt(oCmdAddBooking.Parameters("@HoldingQueueKey").Value.ToString)
            If nHoldingQueueKey > 0 Then
                For Each kvp As KeyValuePair(Of Integer, Integer) In dicOrderItems
                    Try
                        Dim lProductKey As Long = kvp.Key
                        Dim lPickQuantity As Long = kvp.Value
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_ProductCredits_AuthorisationRequestItemAdd", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param53 As SqlParameter = New SqlParameter("@OrderHoldingQueueKey", SqlDbType.Int, 4)
                        param53.Value = nHoldingQueueKey
                        oCmdAddStockItem.Parameters.Add(param53)
                        Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param54.Value = lProductKey
                        oCmdAddStockItem.Parameters.Add(param54)
                        Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                        param56.Value = lPickQuantity
                        oCmdAddStockItem.Parameters.Add(param56)

                        Dim param57 As SqlParameter = New SqlParameter("@Authorised", SqlDbType.Char, 1)
                        param57.Value = "N"
                        oCmdAddStockItem.Parameters.Add(param57)

                        oCmdAddStockItem.Connection = oConn
                        oCmdAddStockItem.ExecuteNonQuery()
                    Catch ex As Exception
                        NotifyException("PlaceOrderOnHold", "Could not add product to holding queue", ex)
                    End Try
                Next
                EmailAuthorisers(guidAuthorisationGUID.ToString)
            Else
                NotifyException("PlaceOrderOnHold", "Internal error - no product selected")
            End If
        Catch ex As Exception
            NotifyException("PlaceOrderOnHold", "Could not add order to holding queue", ex)
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Sub EmailAuthorisers(ByVal guidAuthorisationGUID As String)
        Dim dtAuthorisers As DataTable = ExecuteQueryToDataTable("SELECT UserKey, IsPrimaryAuthoriser FROM ClientData_WU_ProductCreditAuthorisers WHERE IsPrimaryAuthoriser = 1")
        If dtAuthorisers.Rows.Count > 0 Then
            'Dim sOneClickURL As String = "http://my.transworld.eu.com/authorise/authcredit.aspx?guid=" & guidAuthorisationGUID
            Dim sOneClickURL As String = "http://my.transworld.eu.com/wurs/"
            Dim sUser As String = ExecuteQueryToDataTable("SELECT FirstName + ' ' + LastName + ' (' + UserID + ')' FROM UserProfile WHERE [key] = " & Session("UserKey")).Rows(0).Item(0)

            Dim sbMessage As New StringBuilder
            sbMessage.Append("User ")
            sbMessage.Append(sUser)
            sbMessage.Append(" has requested one or more products that require authorisation.")
            sbMessage.Append(Environment.NewLine)
            sbMessage.Append(Environment.NewLine)
            sbMessage.Append("You must authorise or decline this request. Go to the authorisation screen at ")
            sbMessage.Append(sOneClickURL)
            sbMessage.Append(Environment.NewLine)
            sbMessage.Append(Environment.NewLine)
            sbMessage.Append("Thank you.")
            sbMessage.Append(Environment.NewLine)
            sbMessage.Append("")

            Dim sbMessageHTML As New StringBuilder
            sbMessageHTML.Append("<html><head></head><body>")
            sbMessageHTML.Append("User ")
            sbMessageHTML.Append(sUser)
            sbMessageHTML.Append(" has requested one or more products that require authorisation.")
            sbMessageHTML.Append("<br />")
            sbMessageHTML.Append(Environment.NewLine)
            sbMessageHTML.Append("<br />")
            sbMessageHTML.Append(Environment.NewLine)
            sbMessageHTML.Append("You must authorise or decline this request. Click ")
            sbMessageHTML.Append("<a href=""" + sOneClickURL + """>here</a> to go to the authorisation screen.")
            sbMessageHTML.Append("<br />")
            sbMessageHTML.Append(Environment.NewLine)
            sbMessageHTML.Append("Thank you.")
            sbMessageHTML.Append("<br />")
            sbMessageHTML.Append(Environment.NewLine)
            sbMessageHTML.Append("</body></html>")
            sbMessageHTML.Append("")
            
            For Each dr As DataRow In dtAuthorisers.Rows
                Call SendMail("CREDIT_AUTH_REQ", ExecuteQueryToDataTable("SELECT EmailAddr FROM UserProfile WHERE [key] = " & dr("UserKey")).Rows(0).Item(0), "Authorisation Request", sbMessage.ToString, sbMessageHTML.ToString)
            Next
        End If
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

    Protected Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dtCnor As DataTable = ExecuteQueryToDataTable("SELECT * FROM Customer WHERE CustomerKey = " & pnImpersonateCustomer)
        If dtCnor.Rows.Count = 1 Then
            gdrCnor = dtCnor.Rows(0)
        Else
            WebMsgBox.Show("Couldn't find Consignor details.")
            Exit Sub
        End If
        gdtBasket = Session("BO_BasketData")
        If gdtBasket Is Nothing Then
            WebMsgBox.Show("Basket not found.")
            Exit Sub
        End If
        Dim nConsignmentNumber As Int32 = 0
        Dim bPlaceOrder As Boolean = False
        Dim bCreateAuthorisationRequests As Boolean = False
        For Each dr As DataRow In gdtBasket.Rows
            If dr("QtyGranted") > 0 Then
                bPlaceOrder = True
            End If
            If dr("OrderStatus") = ORDER_STATUS_REDUCED_TO_CREDIT_LIMIT_REQUEST_AUTHORISATION Then
                bCreateAuthorisationRequests = True
            End If
        Next
        If bPlaceOrder Then
            nConsignmentNumber = nSubmitConsignment()
            If nConsignmentNumber > 0 Then
                trConsignmentNumber1.Visible = True
                trConsignmentNumber2.Visible = True
                lblConsignmentNumber.Text = nConsignmentNumber.ToString
                gvBasket.DataSource = Nothing
                gvBasket.DataBind()
                btnPlaceOrder.Visible = False
                'lblConsignment.Text = "Consignment # " & nConsignmentNumber
            Else
                mpe.PopupControlID = "divOrderError"
                mpe.Show()
            End If
        Else
            gvBasket.DataSource = Nothing
            gvBasket.DataBind()
            btnPlaceOrder.Visible = False
        End If
        If bCreateAuthorisationRequests Then
            Call ProcessAuthorisationRequests(nConsignmentNumber)
            Call CheckForOverdraftRequests()
        End If
        Call AdjustUserCredits()
        ' need to summarise order, if any, returning consignment number, and if one or more authorisations required say that these have been requested.

        If SHOW_SAVEADDRESS AndAlso cbSaveAddress.Checked Then
            Call UpdateAddressCookie()
        End If
        Call ClearUp()
    End Sub
    
    Protected Sub UpdateAddressCookie()
        
        Dim c As HttpCookie = New HttpCookie(LAST_ADDRESS_COOKIE)
        c.Values.Add("ContactName", tbCneeCtcName.Text)
        c.Values.Add("Name", tbCneeName.Text)
        c.Values.Add("Addr1", tbCneeAddr1.Text)
        c.Values.Add("Addr2", tbCneeAddr2.Text)
        c.Values.Add("Addr3", tbCneeAddr3.Text)
        c.Values.Add("Town", tbCneeTown.Text)
        c.Values.Add("State", tbCneeState.Text)
        c.Values.Add("Postcode", tbCneePostCode.Text)
        c.Values.Add("CountryCode", ddlCountry.SelectedValue)
        c.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(c)
    End Sub
    
    Protected Sub ClearOrderConfirmation()
        lblConsignmentNumber.Text = String.Empty
        trConsignmentNumber1.Visible = False
        trConsignmentNumber2.Visible = False
        'If tbSearch.Text <> String.Empty Then
        '    tbSearch.Text = String.Empty
        '    'Call PopulateProductDropdown(pnImpersonateCustomer)         ' XXXX
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
  
    Protected Sub rcbProduct_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim s As String = e.Text
        Dim data As DataTable = GetProductsByCustomer(pnImpersonateCustomer, e.Text)
        'Dim sThumbnailImage As String = String.Empty
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbProduct.DataTextField = "Product"
        rcbProduct.DataValueField = "LogisticProductKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            'rcb.Text = data.Rows(i)("Product").ToString() + " (max order: " + data.Rows(i)("Quantity").ToString() + ")"
            rcb.Text = data.Rows(i)("Product").ToString()
            rcb.Value = data.Rows(i)("LogisticProductKey").ToString()
            'sThumbnailImage = data.Rows(i)("ThumbnailImage").ToString()
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
        If nOffset <= ITEMS_PER_REQUEST Then
            GetStatusMessage = "++++ Click for more items ++++"
        End If
        If nOffset = nTotal Then
            GetStatusMessage = "No more items"
        Else
            GetStatusMessage = "++++ Click for more items ++++"
        End If
    End Function
    
    Protected Sub rcbProduct_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        If rcb.SelectedIndex = 0 Then
            btnAddToOrder.Enabled = False
            rntbQty.Enabled = False
        Else
            btnAddToOrder.Enabled = True
            rntbQty.Enabled = True
            rntbQty.Focus()
        End If
        Call ClearOrderConfirmation()
    End Sub
    
    Protected Sub lnkbtnPlaceAnotherOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearOrderConfirmation()
        'Call CreateOrRetrieveCookie()
        rcbProduct.Focus()
    End Sub
    
    Protected Sub lnkbtnClearSearchTerm_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearSearchTerm()
    End Sub
    
    Protected Sub ClearSearchTerm()
        rcbProduct.Focus()
    End Sub
    
    Protected Sub lnkbtnRemove_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgbtn As ImageButton = sender
        Dim sLogisticProductKey As String = imgbtn.CommandArgument
        Call RemoveItemFromBasket(sLogisticProductKey)
    End Sub
    
    Protected Sub SetAddressFieldsVisibility(ByVal bVisible As Boolean)
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
        
    Protected Sub SetAddressBookFieldsVisibility(ByVal bVisible As Boolean)
        trAddressBook01.Visible = bVisible
        trAddressBook02.Visible = bVisible
        trAddressBook03.Visible = bVisible
        trAddressBook04.Visible = bVisible
    End Sub

    Protected Sub lnkbtnShowAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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
    
    Protected Sub SetProductSelectability(ByVal bEnabled As Boolean)
        rcbProduct.Enabled = bEnabled
        'cbFilterProducts.Enabled = bEnabled
        'btnSearchGo.Enabled = bEnabled
        'lnkbtnClearSearchTerm.Enabled = bEnabled
    End Sub
    
    Protected Sub btnSelectAddress_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = sender
        Dim nAddressKey As Integer = btn.CommandArgument
        Call SetAddressBookFieldsVisibility(False)
        Call SetAddressFieldsVisibility(True)
        Call GetConsigneeAddress(nAddressKey)
        Call SetPlaceOrderButtonVisibility()
        Call SetProductSelectability(True)
        lnkbtnShowAddressBook.Text = "show address book"
    End Sub
    
    Protected Sub btnSearchAddressBook_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub gvAddressBook_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvAddressBook.PageIndex = e.NewPageIndex
        Call BindAddressBook()
    End Sub
    
    Protected Sub rbPersonalAddressBook_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub

    Protected Sub rbSharedAddressBook_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub GetConsigneeAddress(ByVal nAddressKey As Int32)
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

    Protected Sub lnkbtnClearFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rcbProduct.Text = String.Empty
        btnAddToOrder.Enabled = False
        rcbProduct.Focus()
    End Sub
    
    Protected Sub lnkbtnClearAddressBookSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbSearchAddressBook.Text = String.Empty
        gvAddressBook.PageIndex = 0
        Call BindAddressBook()
    End Sub
    
    Protected Sub rbProductGroup_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.ID = "rbAllProducts" Then
            trPopularProductsRotator.Visible = False
            trPopularProductsHelpBar.Visible = False
        Else
            trPopularProductsRotator.Visible = True
            trPopularProductsHelpBar.Visible = True
        End If
        rbAllProducts.ForeColor = System.Drawing.Color.FromName(System.Drawing.KnownColor.WindowText)
        rbAllProducts.Font.Bold = False
        rbFavouriteProducts.ForeColor = System.Drawing.Color.FromName(System.Drawing.KnownColor.WindowText)
        rbFavouriteProducts.Font.Bold = False
        If rb.Checked Then
            rb.ForeColor = Drawing.Color.Blue
            rb.Font.Bold = True
        End If
        rcbProduct.Text = String.Empty
        rcbProduct.Focus()
    End Sub
    
    Protected Sub btnRedirect_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlMain.Visible = True
        pnlRedirect.Visible = False
    End Sub
    
    Protected Sub imgbtnHelp_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        imgbtnHelp.Visible = False
    End Sub
    
    Protected Sub gvBasket_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            'Dim hidCreditStatus As HiddenField
            'hidCreditStatus = gvr.Cells(2).FindControl("hidCreditStatus")
            'Dim arrCreditStatus() As String = hidCreditStatus.Value.Split(",")
            'Dim nRemainingCredit As Int32 = arrCreditStatus(0)
            'Dim nProductCreditStatus As Int32 = arrCreditStatus(1)
            'Dim sRefreshDate As Int32 = arrCreditStatus(2)
        End If
    End Sub
    
    Protected Sub imgbtnAvailable_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgbtn As ImageButton = sender
    End Sub

    Protected Function sProductAvailable(ByVal DataItem As Object) As String
        sProductAvailable = String.Empty
        'Dim nAvailable As Int32 = DataBinder.Eval(DataItem, "Available") ' here
        Dim nAvailable As Int32 = DataBinder.Eval(DataItem, "QtyGranted") ' here
        sProductAvailable = nAvailable
    End Function

    Protected Function sProductCreditStatus(ByVal DataItem As Object) As String
        sProductCreditStatus = DataBinder.Eval(DataItem, "Message")
        Exit Function
        sProductCreditStatus = String.Empty
        Dim sMessage As String = String.Empty
        'Dim nAvailable As Int32 = DataBinder.Eval(DataItem, "Available") ' here
        Dim nAvailable As Int32 = DataBinder.Eval(DataItem, "QtyGranted") ' here
        Dim nRemainingCredit As Int32 = DataBinder.Eval(DataItem, "RemainingCredit")
        Dim nProductCreditStatus As Int32 = DataBinder.Eval(DataItem, "EnforceCreditLimit")
        Dim sCreditStatus As String
        If DataBinder.Eval(DataItem, "EnforceCreditLimit") = 0 Then
            sCreditStatus = "enforce credit limit"
        Else
            sCreditStatus = "more on request"
        End If
        Dim sRefreshDate As String = DataBinder.Eval(DataItem, "RefreshDate")
        If sRefreshDate = String.Empty Then
            sRefreshDate = " (no refresh date available)"
        End If
        sProductCreditStatus = "Remaining credit: " & nRemainingCredit.ToString & "; " & sCreditStatus & "; " & sRefreshDate
    End Function

    'Protected Sub RadToolTipManager_AjaxUpdate(ByVal sender As Object, ByVal e As ToolTipUpdateEventArgs)
    '    If e.Value.Contains(",") Then
    '        If e.TargetControlID.Contains("lnkConsignment") Then
    '            'Call BindStockItems(e)
    '            'Call GetTracking(e)
    '        ElseIf e.TargetControlID.Contains("lnkTopicReference") Then
    '            'Call GetUserData(e)
    '        End If
    '    End If
    'End Sub

    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hm As New HtmlMeta
        hm.HttpEquiv = "X-UA-Compatible"
        hm.Content = "IE=9"
        Page.Header.Controls.AddAt(0, hm)
    End Sub
    
    Protected Function sGetOverdraftDetails(ByVal DataItem As Object) As String
        Dim sSQL As String
        Dim sbOverdraftDetails As New StringBuilder
        Dim nID As Int32 = DataBinder.Eval(DataItem, "id")
        sSQL = "SELECT [id], ISNULL(COnsignmentKey, 0) 'ConsignmentKey' FROM ProductCreditsOrderHoldingQueue WHERE [id] = " & nID
        Dim drOverdraftRequest As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        Dim dtOverdraftItems As DataTable = ExecuteQueryToDataTable("SELECT * FROM ProductCreditsOrderItemHoldingQueue WHERE ProductCreditsOrderHoldingQueueKey = " & drOverdraftRequest("id") & " ORDER BY [id]")
        For Each drOverdraftItem As DataRow In dtOverdraftItems.Rows
            Dim sProductDetails() As String = GetProductDetailsFromProductKey(drOverdraftItem("LogisticProductKey"))
            sbOverdraftDetails.Append(sProductDetails(0))
            sbOverdraftDetails.Append(" ")
            sbOverdraftDetails.Append(sProductDetails(1))
            sbOverdraftDetails.Append(" ")
            sbOverdraftDetails.Append(sProductDetails(2))
            sbOverdraftDetails.Append(" ")
            sbOverdraftDetails.Append("Qty: ")
            'sbOverdraftDetails.Append(GetProductDetailsFromProductKey(drOverdraftItem("LogisticProductKey")))
            'sbOverdraftDetails.Append(" ")
            sbOverdraftDetails.Append(drOverdraftItem("ItemsOut"))
            sbOverdraftDetails.Append(Environment.NewLine)
        Next
        Dim nConsignmentKey As Int32 = drOverdraftRequest("ConsignmentKey")
        If nConsignmentKey > 0 Then
            sbOverdraftDetails.Append(" (Consignment: ")
            sbOverdraftDetails.Append(nConsignmentKey.ToString)
            sbOverdraftDetails.Append(")")
            sbOverdraftDetails.Append(Environment.NewLine)
        End If
        sGetOverdraftDetails = sbOverdraftDetails.ToString
    End Function
    
    Protected Sub lnkbtnHideOverdraftRequests_Click(sender As Object, e As System.EventArgs)
        Call SetAuthRequestsVisibility(False)
    End Sub
    
    Protected Sub gvOverdraftRequests_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidOrderStatus As HiddenField = gvr.Cells(0).FindControl("hidOrderStatus")
            If hidOrderStatus.Value.ToLower = "declined" Then
                gvr.BackColor = Drawing.Color.OrangeRed
            ElseIf hidOrderStatus.Value.ToLower = "authorised" Then
                gvr.BackColor = Drawing.Color.PaleGreen
            End If
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Quick Order</title>
    <style type="text/css">
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
        .style1
        {
            width: 60px;
        }
        .style2
        {
            width: 60px;
            height: 16px;
        }
        .style3
        {
            height: 16px;
        }
        .style4
        {
            height: 33px;
        }
    </style>
    <link href="css/modalpopup.css" rel="stylesheet" type="text/css" />
</head>
<script type="text/javascript">

    function ImageClick(productKey) {

        PageMethods.SetCustomTextOnProductDropDown(productKey, OnSuccess);
    }


    function OnSuccess(result) {
        var rcb = $find('<%= rcbProduct.ClientID %>');
        rcb.set_text(result);
    } 
      
</script>
<%--<script type="text/javascript">
    function showToolTip(element, id) {
        var tooltipManager = $find("<%= rttm.ClientID %>");
        //If the user hovers the image before the page has loaded, there is no manager created
        if (!tooltipManager) return;
        //Find the tooltip for this element if it has been created 
        var tooltip = tooltipManager.getToolTipByElement(element);
        //Create a tooltip if no tooltip exists for such element 
        if (!tooltip) {
            tooltip = tooltipManager.createToolTip(element);
        }
        //Let the tooltip's own show mechanism take over from here - execute the onmouseover just once
        element.onmouseover = null;
        //show the tooltip
        setTimeout(function () {
            tooltip.set_value(element.innerText + "," + id);
            tooltip.show();
        }, 10);
    }
</script>
--%><body style="font-size: xx-small; font-family: Verdana"><%-- style="font-size: 9pt; --%><form
id="frmOrder" runat="server">
<%--<asp:ScriptManager runat="server" />--%>
<%--<asp:ScriptManager runat="server" EnablePageMethods="true" />--%>
<main:Header ID="ctlHeader" runat="server" />
<asp:PlaceHolder ID="PlaceHolderForScriptManager" runat="server"/>
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
<asp:Panel ID="pnlMain" runat="server" Width="100%">
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
                        AutoPostBack="True" Font-Size="X-Small">
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
            <td colspan="2">
                <table width="100%">
                    <tr id="trPopularProductsRotator" runat="server" visible="false">
                        <td class="style1">
                            &nbsp;
                        </td>
                        <td>
                            <telerik:RadRotator ID="radRotator" Width="527" ScrollDirection="Right" RotatorType="AutomaticAdvance"
                                ScrollDuration="2000" FrameDuration="2000" runat="server" Height="104px">
                                <ItemTemplate>
                                    <div>
                                        <asp:Image ID="imgRotator" runat="server" ImageAlign="Middle" Style="margin: 2px;
                                            border: 1px solid" ImageUrl='<%# GetImage(Container.DataItem) %>' Height="100px"
                                            Width="100px" />
                                        <asp:HiddenField ID="hidLogisticProductKey" Value='<%# Bind("LogisticProductKey") %>'
                                            runat="server" />
                                        <telerik:RadToolTip ID="rttp" runat="server" TargetControlID="imgRotator" RelativeTo="Element"
                                            Position="TopRight" RenderInPageRoot="true">
                                            <asp:Label ID="lblProductDescriptionTooltip" Text='<%# Bind("Product") %>' runat="server" />
                                        </telerik:RadToolTip>
                                    </div>
                                </ItemTemplate>
                            </telerik:RadRotator>
                        </td>
                    </tr>
                    <tr>
                        <td class="style2">
                        </td>
                        <td class="style3">
                            <table>
                                <tr id="trPopularProductsHelpBar" runat="server" visible="false">
                                    <td style="width: 350px">
                                        <asp:ImageButton ID="imgbtnHelp" runat="server" ImageUrl="~/images/WUsearchbarhelp.gif"
                                            OnClick="imgbtnHelp_Click" CausesValidation="False" ToolTip="click to remove these messages" />
                                        &nbsp;
                                    </td>
                                    <td>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr valign="top">
            <td style="width: 50%">
                <table>
                    <tr>
                        <td style="width: 65px" />
                        <td style="width: 550px">
                            <asp:Label ID="Label12" runat="server" Text="Products" Font-Bold="True" Font-Size="Small" />
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:RadioButton ID="rbAllProducts" runat="server" GroupName="Products" Checked="True"
                                Text="All products" AutoPostBack="True" OnCheckedChanged="rbProductGroup_CheckedChanged"
                                ForeColor="Blue" Font-Bold="True" ToolTip="include ALL available products in the search bar" />
                            &nbsp;
                            <asp:RadioButton ID="rbFavouriteProducts" runat="server" GroupName="Products" Text="Most popular products"
                                OnCheckedChanged="rbProductGroup_CheckedChanged" AutoPostBack="True" ToolTip="include only the most popular products (if available) in the search bar" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                        <td>
                            <table style="width: 100%">
                                <tr>
                                    <td style="width: 85%">
                                        <telerik:RadComboBox ID="rcbProduct" runat="server" Width="100%" Font-Names="Arial"
                                            Font-Size="X-Small" Font-Bold="true" OnSelectedIndexChanged="rcbProduct_SelectedIndexChanged"
                                            AutoPostBack="True" HighlightTemplatedItems="true" CausesValidation="False" EnableLoadOnDemand="True"
                                            OnItemsRequested="rcbProduct_ItemsRequested" EnableVirtualScrolling="True" ShowMoreResultsBox="True"
                                            Filter="Contains" EmptyMessage="Click here to show products, or type a search phrase"
                                            ToolTip="Shows all available products when no search text is specified. Search for products by typing a product code or description.">
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
                                    <td style="width: 15%" align="right">
                                        &nbsp;<asp:LinkButton ID="lnkbtnClearFilter" runat="server" OnClick="lnkbtnClearFilter_Click"
                                            CausesValidation="False" ToolTip="Clears product search text, if present." Font-Size="XX-Small">clear search</asp:LinkButton>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                        </td>
                        <td>
                            <asp:Label ID="Label560" runat="server" Text="Quantity:" Style="font-size: 9pt" />
                            &nbsp;<telerik:RadNumericTextBox ID="rntbQty" runat="server" Font-Bold="True" Font-Size="X-Small"
                                MaxValue="100000" MinValue="1" ShowSpinButtons="True" Value="1" Width="50px">
                                <NumberFormat DecimalDigits="0" />
                            </telerik:RadNumericTextBox>
                            &nbsp;<asp:Button ID="btnAddToOrder" runat="server" Text="Add to Basket" Width="169px"
                                OnClick="btnAddToOrder_Click" CausesValidation="False" Enabled="False" ToolTip="add the product in the search bar to your basket" />
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
                            &nbsp;
                        </td>
                        <td>
                            <asp:Label ID="Label557" runat="server" Font-Bold="True" Font-Size="Small" Text="Basket" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:GridView ID="gvBasket" runat="server" AutoGenerateColumns="False" CellPadding="2"
                                Width="100%" OnRowDataBound="gvBasket_RowDataBound">
                                <AlternatingRowStyle BackColor="#FFFF99" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgBtnRemove" runat="server" CausesValidation="false" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>'
                                                ImageUrl="~/images/delete.gif" OnClick="lnkbtnRemove_Click" />
                                            <asp:HiddenField ID="hidRequiresAuthorisation" runat="server" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Product" HeaderText="Product" ReadOnly="True" />
                                    <asp:BoundField DataField="QtyRequested" HeaderText="Requested Qty" ReadOnly="True"
                                        SortExpression="QtyRequested" />
                                    <asp:TemplateField HeaderText="Available Qty">
                                        <ItemTemplate>
                                            <%--<asp:Label ID="lblAvailable" runat="server" Text='<%# Bind("Available") %>'/>--%>
                                            <asp:Label ID="Label2" runat="server" Text='<%# sProductAvailable(Container.DataItem) %>' />
                                            <asp:ImageButton ID="imgbtnAvailable" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>'
                                                CausesValidation="False" OnClick="imgbtnAvailable_Click" Visible='<%# sProductCreditStatus(Container.DataItem) <> String.Empty %>'
                                                ImageUrl="~/images/information-icon.png" />
                                            <telerik:RadToolTip ID="RadToolTipStatus" runat="server" TargetControlID="imgbtnAvailable"
                                                Text='<%# sProductCreditStatus(Container.DataItem) %>' ShowEvent="OnClick" HideEvent="LeaveTargetAndToolTip"
                                                Width="300px" AutoCloseDelay="6000">
                                                (tooltip text not set)</telerik:RadToolTip>
                                        </ItemTemplate>
                                        <ItemStyle ForeColor="#999999" />
                                    </asp:TemplateField>
                                </Columns>
                                <EmptyDataRowStyle BackColor="#FFFFCC" />
                                <EmptyDataTemplate>
                                    Your basket is empty
                                </EmptyDataTemplate>
                                <RowStyle BackColor="#FFFFCC" />
                            </asp:GridView>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:Button ID="btnPlaceOrder" runat="server" OnClick="btnPlaceOrder_Click" Text="Place Order"
                                ToolTip="check the order, request confirmation, then submit it" Visible="False"
                                Width="150px" />
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
                    <tr id="trConsignmentNumber1" runat="server" visible="false">
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:Label ID="lblOrderPlaced" runat="server" Font-Size="X-Small" Text="Thank you for your order. Please note your consignment number:" />
                        </td>
                    </tr>
                    <tr id="trConsignmentNumber2" runat="server" visible="false">
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:Label ID="lblConsignmentNumber" runat="server" Font-Bold="True" Font-Size="Small"
                                ForeColor="Blue" />
                            &nbsp;&nbsp;
                            <asp:LinkButton ID="lnkbtnPlaceAnotherOrder" runat="server" CausesValidation="False"
                                OnClick="lnkbtnPlaceAnotherOrder_Click">place another order</asp:LinkButton>
                        </td>
                    </tr>
                    <tr runat="server" visible="false" id="trAuthRequests01">
                        <td class="style3">
                        </td>
                        <td class="style3">
                        </td>
                    </tr>
                    <tr runat="server" visible="false" id="trAuthRequests02">
                        <td class="style4">
                            &nbsp;
                        </td>
                        <td class="style4">
                            <hr />
                            <asp:Label ID="Label576" runat="server" Style="font-size: 08pt" Text="Overdraft requests (last 30 days):" />
                            &nbsp;<asp:LinkButton ID="lnkbtnHideOverdraftRequests" runat="server" Font-Names="Verdana"
                                Font-Size="XX-Small" OnClick="lnkbtnHideOverdraftRequests_Click" CausesValidation="False">hide</asp:LinkButton>
                        </td>
                    </tr>
                    <tr runat="server" visible="false" id="trAuthRequests03">
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:GridView ID="gvOverdraftRequests" runat="server" CellPadding="2" Width="100%"
                                AutoGenerateColumns="False" OnRowDataBound="gvOverdraftRequests_RowDataBound">
                                <Columns>
                                    <asp:BoundField DataField="OrderCreatedDateTime" HeaderText="Date" ReadOnly="True"
                                        SortExpression="OrderCreatedDateTime" />
                                    <asp:TemplateField HeaderText="Details" SortExpression="id">
                                        <ItemTemplate>
                                            <asp:Label ID="lblOverdraftDetails" runat="server" Text='<%# sGetOverdraftDetails(Container.DataItem) %>' />
                                            <asp:HiddenField ID="hidOrderStatus" Value='<%# Container.DataItem("OrderStatus")%>' runat="server" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
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
                        <td>
                            &nbsp;
                        </td>
                        <td>
                            <asp:Label ID="Label4a1" runat="server" Font-Bold="True" Font-Size="Small" Text="For help using the Quick Order facility, click " />
                            <asp:HyperLink ID="hlinkClickHere" runat="server" Font-Bold="True" Font-Names="Verdana"
                                Font-Size="Small" NavigateUrl="http://my.transworld.eu.com/common/wuhelp/WUQuickOrderHelp.htm"
                                Target="_blank">here</asp:HyperLink>
                        </td>
                    </tr>
                </table>
            </td>
            <td style="width: 50%">
                <table>
                    <tr>
                        <td style="width: 140px" />
                        <td style="width: 3px" />
                        <td style="width: 300px" />
                        <asp:Label ID="Label555" runat="server" Text="Delivery Address" Font-Bold="True"
                            Font-Size="Small" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:LinkButton ID="lnkbtnShowAddressBook" runat="server" OnClick="lnkbtnShowAddressBook_Click"
                            CausesValidation="False" Font-Names="Arial" Font-Size="XX-Small">show address book</asp:LinkButton>
                        <asp:CheckBox ID="cbSaveAddress" runat="server" Checked="True" Text="save address" />
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
                                Text="Shared" Checked="True" OnCheckedChanged="rbSharedAddressBook_CheckedChanged"
                                AutoPostBack="True" />
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
                            <asp:LinkButton ID="lnkbtnClearAddressBookSearch" runat="server" CausesValidation="False"
                                Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnClearAddressBookSearch_Click">clear search</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trAddressBook03" runat="server" visible="false">
                        <td align="right" colspan="3">
                            <asp:GridView ID="gvAddressBook" runat="server" Width="100%" CellPadding="2" Font-Names="Arial"
                                Font-Size="X-Small" AllowPaging="True" AutoGenerateColumns="False" OnPageIndexChanging="gvAddressBook_PageIndexChanging">
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
                                Operator="NotEqual" ValueToCompare="0" Text="###" Style="font-size: 08pt" />
                            &nbsp;<asp:Label ID="Label6" runat="server" Text="Country:" ForeColor="Red" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:DropDownList ID="ddlCountry" runat="server" Width="80%" Style="font-size: 08pt"
                                OnSelectedIndexChanged="ddlCountry_SelectedIndexChanged" AutoPostBack="True" />
                            &nbsp;<asp:LinkButton ID="lnkbtnUK" runat="server" OnClick="lnkbtnUK_Click" CausesValidation="False">UK</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trPostCode" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneePostCode" runat="server" ErrorMessage="###"
                                Style="font-size: 08pt" ControlToValidate="tbCneePostCode" />
                            &nbsp;<asp:Label ID="lblLegendPostcode" runat="server" Text="Postcode:" ForeColor="Red"
                                Style="font-size: 08pt" />
                        </td>
                        <td align="left">
                        </td>
                        <td>
                            <asp:TextBox ID="tbCneePostCode" runat="server" Width="50%" MaxLength="50" Style="font-size: 08pt" />
                            &nbsp;<asp:Button ID="btnPostcodeFind" runat="server" Text="Find" OnClick="btnPostcodeFind_Click"
                                CausesValidation="False" />
                            <asp:Label ID="lblLookupError" runat="server" Visible="False" ForeColor="Red" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trPostcodeLookupOutput" runat="server" visible="false">
                        <td align="right">
                            &nbsp;<asp:Label ID="Label575" runat="server" Style="font-size: 08pt" Text="Select a destination:" />
                        </td>
                        <td />
                        <td>
                            <asp:ListBox ID="lbLookupResults" runat="server" Rows="10" Width="100%" OnSelectedIndexChanged="lbLookupResults_SelectedIndexChanged"
                                AutoPostBack="True" Style="font-size: 08pt" />
                            <br />
                            <asp:LinkButton ID="lnkbtnCancelPostcodeLookup" runat="server" OnClick="lnkbtnCancelPostcodeLookup_Click"
                                CausesValidation="False">cancel</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trContactName" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneeCtcName" runat="server" ControlToValidate="tbCneeCtcName"
                                Enabled="True" ErrorMessage="###" Style="font-size: 08pt" Font-Bold="True" ForeColor="Red" />
                            &nbsp;<asp:Label ID="lblLegendContactName" runat="server" Text="Contact Name:" ForeColor="Red"
                                Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeCtcName" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeName" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneeName" runat="server" ErrorMessage="###" Style="font-size: 08pt"
                                ControlToValidate="tbCneeName" />
                            &nbsp;<asp:Label ID="Label5" runat="server" Text="Name:" ForeColor="Red" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeName" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr1" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvCneeAddr1" runat="server" ErrorMessage="###" Style="font-size: 08pt"
                                ControlToValidate="tbCneeAddr1" />
                            &nbsp;<asp:Label ID="Label7" runat="server" Text="Addr 1:" ForeColor="Red" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr1" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr2" runat="server">
                        <td align="right">
                            <asp:Label ID="Label9" runat="server" Text="Addr 2:" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr2" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr3" runat="server">
                        <td align="right">
                            <asp:Label ID="Label10" runat="server" Text="Addr 3:" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeAddr3" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trTownCity" runat="server">
                        <td align="right">
                            <asp:RequiredFieldValidator ID="rfvTownCity" runat="server" ErrorMessage="###" Style="font-size: 08pt"
                                ControlToValidate="tbCneeTown" />
                            &nbsp;<asp:Label ID="Label11" runat="server" Text="Town/City:" ForeColor="Red" Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td class="style4">
                            <asp:TextBox ID="tbCneeTown" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeState" runat="server">
                        <td align="right">
                            <asp:Label ID="lblLegendCountyStateRegionProvince" runat="server" Text="County / Region:"
                                Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td class="style4">
                            <asp:TextBox ID="tbCneeState" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trUSStatesCanadianProvinces" runat="server" visible="false">
                        <td align="right">
                            &nbsp;<asp:CompareValidator ID="cvStateProvince" runat="server" ControlToValidate="ddlUSStatesCanadianProvinces"
                                Operator="NotEqual" ValueToCompare="0" Text="###" Style="font-size: 08pt" />
                            &nbsp;<asp:Label ID="lblLegendStateProvince" runat="server" Text="State:" ForeColor="Red"
                                Style="font-size: 08pt" />
                        </td>
                        <td />
                        <td>
                            <asp:DropDownList ID="ddlUSStatesCanadianProvinces" runat="server" Width="100%" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trUSStateShortcuts" runat="server" visible="false">
                        <td align="right">
                            &nbsp;
                        </td>
                        <td />
                        <td>
                            <asp:LinkButton ID="lnkbtnNewYorkCity" runat="server" Font-Names="Verdana" Style="font-size: 08pt"
                                OnClick="lnkbtnNewYorkCity_Click" CausesValidation="False">NYC</asp:LinkButton>
                            &nbsp;<asp:LinkButton ID="lnkbtnWashingtonDC" runat="server" Font-Names="Verdana"
                                Style="font-size: 08pt" OnClick="lnkbtnWashingtonDC_Click" CausesValidation="False">Washington D.C.</asp:LinkButton>
                        </td>
                    </tr>
                    <tr id="trCneeNameReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label565" runat="server" Style="font-size: 08pt" Text="Name:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeNameReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr1ReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label566" runat="server" Style="font-size: 08pt" Text="Addr 1:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeAddr1ReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr2ReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label567" runat="server" Style="font-size: 08pt" Text="Addr 2:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeAddr2ReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeAddr3ReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label568" runat="server" Style="font-size: 08pt" Text="Addr 3:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeAddr3ReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeTownCityReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label569" runat="server" Style="font-size: 08pt" Text="Town / City:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeTownCityReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeStateReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label570" runat="server" Style="font-size: 08pt" Text="County / Region:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneeStateReadOnly" runat="server" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneePostcodeReadOnly" runat="server" visible="false">
                        <td align="right">
                            <asp:Label ID="Label571" runat="server" Style="font-size: 08pt" Text="Postcode:" />
                        </td>
                        <td />
                        <td>
                            <asp:Label ID="lblCneePostcodeReadOnly" runat="server" Style="font-size: 08pt" />
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
                            <asp:Label ID="Label572" runat="server" Style="font-size: 08pt" Text="Contact Tel:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeTel" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCneeEmail" runat="server">
                        <td align="right">
                            <asp:Label ID="Label573" runat="server" Style="font-size: 08pt" Text="Contact Email:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCneeEmail" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCustRef1" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef1" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef1" Style="font-size: 08pt">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef1" runat="server" Style="font-size: 08pt" Text="Cust Ref 1:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef1" runat="server" Width="100%" MaxLength="25" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCustRef2" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef2" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef2" Style="font-size: 08pt">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef2" runat="server" Style="font-size: 08pt" Text="Cust Ref 2:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef2" runat="server" Width="100%" MaxLength="25" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCustRef3" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef3" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef3" Style="font-size: 08pt">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef3" runat="server" Style="font-size: 08pt" Text="Cust Ref 3:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef3" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCustRef4" runat="server">
                        <td align="right">
                            &nbsp;<asp:RequiredFieldValidator ID="rfvCustRef4" runat="server" Enabled="false"
                                EnableClientScript="false" ControlToValidate="tbCustRef4" Style="font-size: 08pt">###</asp:RequiredFieldValidator>
                            &nbsp;<asp:Label ID="lblLegendCustRef4" runat="server" Style="font-size: 08pt" Text="Cust Ref 4:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbCustRef4" runat="server" Width="100%" MaxLength="50" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            <asp:Label ID="Label574" runat="server" Style="font-size: 08pt" Text="Special Instructions:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbSpecialInstructions" runat="server" Width="100%" MaxLength="500"
                                Rows="3" TextMode="MultiLine" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trPackingNote" runat="server">
                        <td align="right">
                            <asp:Label ID="Label3" runat="server" Style="font-size: 08pt" Text="Packing Note:" />
                        </td>
                        <td />
                        <td>
                            <asp:TextBox ID="tbPackingNote" runat="server" Width="100%" MaxLength="50" Rows="3"
                                TextMode="MultiLine" Style="font-size: 08pt" />
                        </td>
                    </tr>
                    <tr id="trCheckAddressXX" runat="server">
                        <td align="right" colspan="3">
                            &nbsp;
                        </td>
                    </tr>
                    <tr id="trCheckAddress" runat="server">
                        <td align="right" colspan="3">
                            <asp:Label ID="lblLegendCheckAddress" runat="server" Font-Bold="True" ForeColor="Maroon"
                                Text="Please check the address we have on file for you, as shown above, is correct." />
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
            <asp:Label ID="Label559" runat="server" Font-Size="Small" Text="Are you sure you want to submit this order?" />
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
            <asp:Label ID="Label8" runat="server" Font-Size="Small" Text="Sorry, we are unable to complete your order.<br /><br />This can occur if you refresh your page immediately after placing an order.<br /><br />For further assistance please contact Transworld Customer Services (0207 231 3131)." />
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
    --%>
    <%--    <ajaxToolkit:TextBoxWatermarkExtender ID="TBWEQty" runat="server" TargetControlID="tbQty"
        WatermarkText=" " WatermarkCssClass="watermarked" />
    --%>
</asp:Panel>
<asp:Panel ID="pnlRedirect" runat="server" Width="100%" Visible="False">
    <p style="text-align: center;">
        You have new messages.</p>
    <p style="text-align: center;">
        Please read your messages before placing further orders.</p>
    <p style="text-align: center;">
        Thank you.</p>
    <p style="text-align: center;">
        <asp:Button ID="btnRedirect" runat="server" Text="OK" Width="100px" OnClick="btnRedirect_Click" />
    </p>
</asp:Panel>
    </form>
</body>
</html>
