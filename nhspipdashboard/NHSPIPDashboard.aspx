<%@ Page Language="VB" Theme="AIMSDefault" MaintainScrollPositionOnPostback="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.Drawing.Image" %>
<%@ Import Namespace="System.Drawing.Color" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="System.Threading" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Net" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
    
    ' Sort out virtual orgs only showing when show PCTs is selected
    
    Const NHSPIP_CUSTOMER_KEY As Integer = 580
    Const NHSPIPTEST_CUSTOMER_KEY As Integer = 16

    Const LOG_ENTRY_TYPE_ORDER As String = "ORDER"
    Const LOG_ENTRY_TYPE_PRODUCT As String = "PRODUCT"
    Const LOG_ENTRY_TYPE_ACCOUNT As String = "ACCOUNT"

    Const NHSPIP_WEBFORM_CONTROL_USERID As String = "NHSPIPPCTWebform"
    Const NHSPIPTEST_WEBFORM_CONTROL_USERID As String = "NHSPIPTESTPCTWebform"
    
    Dim gsWebFormControlUserId As String = NHSPIP_WEBFORM_CONTROL_USERID
    Dim gnCustomerKey As Integer = NHSPIP_CUSTOMER_KEY
    Dim gsSQL As String
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call HideAllPanels()
            pnlIntro.Visible = True
        End If
    End Sub

    Protected Sub HideAllPanels()
        pnlIntro.Visible = False
        pnlGeneric.Visible = False
        pnlAssignedByOrganisation.Visible = False
        pnlAssignedByCreationDate.Visible = False
        pnlAwaitingAllocation.Visible = False
        pnlReservations.Visible = False
        pnlOrders.Visible = False
        pnlBackOrders.Visible = False
        pnlBackOrdersData.Visible = False
        pnlGoodsIn.Visible = False
        pnlAvailableForDistribution.Visible = False
        pnlDistribute.Visible = False
        pnlActivity.Visible = False
        pnlNoProducts.Visible = False
        pnlNoAccounts.Visible = False
        pnlSearchOrgs.Visible = False
        pnlOrganisationList.Visible = False
        pnlProductVisibility.Visible = False
        pnlAccounts.Visible = False
        pnlConsistencyCheck.Visible = False
        pnlHelp.Visible = False
        pnlDownload.Visible = False
        pnlData.Visible = False
        pnlData2.Visible = False
        pnlAddNote.Visible = False
    End Sub

    Protected Sub Menu1_MenuItemClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MenuEventArgs)
        Call HideAllPanels()
        lblVisibilityMismatches.Visible = False
        btnVisibilityMismatches.Visible = False
        lblLegendMismatches.Visible = False
        gvData.EmptyDataText = "No data retrieved"
        gvData.Width = New Unit("100%")
        Select Case Menu1.SelectedValue.ToLower
            Case "generic"
                Call DisplayGeneric()
                pnlGeneric.Visible = True
                pnlData.Visible = True
            Case "assignedbyorganisation"
                Call DisplayAssignedByOrganisation()
                pnlAssignedByOrganisation.Visible = True
                pnlData.Visible = True
            Case "assignedbycreationdate"
                Call DisplayAssignedbyCreationDate()
                pnlAssignedByCreationDate.Visible = True
                pnlData.Visible = True
            Case "awaitingallocation"
                Call DisplayAwaitingAllocation()
                pnlAwaitingAllocation.Visible = True
                pnlData.Visible = True
            Case "reservations"
                Call DisplayReservations()
                pnlReservations.Visible = True
            Case "orders"
                Call DisplayOrders()
                pnlOrders.Visible = True
            Case "backorders"
                Call DisplayBackOrders()
                pnlBackOrders.Visible = True
                pnlBackOrdersData.Visible = True
            Case "goodsin"
                Call DisplayGoodsIn()
                pnlGoodsIn.Visible = True
            Case "availablefordistribution"
                Call DisplayAvailableForDistribution()
                pnlAvailableForDistribution.Visible = True
            Case "activity"
                Call DisplayActivity()
                pnlActivity.Visible = True
            Case "noproducts"
                Call NoProducts()
                pnlNoProducts.Visible = True
                pnlData.Visible = True
            Case "noaccounts"
                Call NoAccounts()
                pnlNoAccounts.Visible = True
                pnlData.Visible = True
            Case "searchorgs"
                gsSQL = String.Empty
                pnlSearchOrgs.Visible = True
                ' tbSearchOrgs.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchGo.ClientID + "')")
                pnlSearchOrgs.DefaultButton = "btnSearchGo"
                tbSearchOrgs.Focus()
            Case "organisationlist"
                Call DisplayOrganisations()
                pnlOrganisationList.Visible = True
                pnlData.Visible = True
            Case "productvisibility"
                Call ProductVisibility()
                pnlProductVisibility.Visible = True
                pnlData.Visible = True
            Case "accounts"
                Call Accounts()
                pnlAccounts.Visible = True
                pnlData.Visible = True
            Case "consistencycheck"
                Call ConsistencyCheck()
                pnlConsistencyCheck.Visible = True
                pnlData.Visible = True
                pnlData2.Visible = True
            Case "changehistory"
                pnlIntro.Visible = True
            Case "help"
                pnlHelp.Visible = True
            Case Else
        End Select
        psDownloadQuery = gsSQL
        If psDownloadQuery <> String.Empty Then
            pnlDownload.Visible = True
        Else
            pnlDownload.Visible = False
        End If
    End Sub

    Protected Sub DisplayReservations()
        Call SetStockReservationControlEnable(False)
        Call ClearReservationsControls()
        Call InitStockReservationOrganisations()
    End Sub
    
    Protected Sub DisplayOrders()
        rbOrdersAll.Checked = True
    End Sub
    
    Protected Sub DisplayBackOrders()
        Dim sbSQL1 As New StringBuilder
        sbSQL1.Append("DECLARE @x int ")
        sbSQL1.Append("DECLARE @Total int ")
        sbSQL1.Append("DECLARE @ProductCode varchar(50) ")
        sbSQL1.Append("DECLARE @ProductDescription varchar(200) ")
        sbSQL1.Append("DECLARE @Note varchar(4000) ")
        sbSQL1.Append("CREATE TABLE #temp (LogisticProductKey int, ProductCode varchar(50), ProductDescription varchar(200), Qty int, Note varchar(8000)) ")
        sbSQL1.Append("DECLARE c CURSOR FOR SELECT DISTINCT GenericProductKey FROM NHSPIPLinkedProducts ")
        sbSQL1.Append("OPEN c ")
        sbSQL1.Append("FETCH NEXT FROM c INTO @x ")
        sbSQL1.Append("WHILE (@@FETCH_STATUS) = 0 ")
        sbSQL1.Append("BEGIN ")
        sbSQL1.Append("SELECT @Total = SUM(BacklogQty) FROM NHSPIPLinkedProducts WHERE GenericProductKey = @x ")
        sbSQL1.Append("SET @ProductCode = (SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = @x) ")
        sbSQL1.Append("SET @ProductDescription = (SELECT ProductDescription FROM LogisticProduct WHERE LogisticProductKey = @x) ")
        sbSQL1.Append("SET @Note = (SELECT NoteText FROM NHSPIPBackOrderNote WHERE LogisticProductKey = @x) ")
        sbSQL1.Append("IF @Total > 0 ")
        sbSQL1.Append("INSERT INTO #temp (LogisticProductKey, ProductCode, ProductDescription, Qty, Note) ")
        sbSQL1.Append("VALUES (@x, @ProductCode, @ProductDescription, @Total,  ISNULL(@Note,'')) ")
        sbSQL1.Append("FETCH NEXT FROM c INTO @x ")
        sbSQL1.Append("END ")
        sbSQL1.Append("CLOSE c ")
        sbSQL1.Append("DEALLOCATE c ")
        sbSQL1.Append("SELECT * FROM #temp ORDER BY ProductCode ")
        sbSQL1.Append("DROP TABLE #temp ")
        
        Dim sSQL2 As String = "SELECT nplp.RingFencedProductKey 'LogisticProductKey', lp.ProductCode 'Product', lp.ProductDescription 'Description', lp.ProductDate 'Organisation', BacklogQty 'Qty', npbon.NoteText 'Note' FROM NHSPIPLinkedProducts nplp INNER JOIN LogisticProduct lp on nplp.RingFencedProductKey = lp.LogisticProductKey LEFT OUTER JOIN NHSPIPBackOrderNote npbon ON nplp.RingFencedProductKey = npbon.LogisticproductKey WHERE BacklogQty > 0"
        Dim oDataTable1 As DataTable = ExecuteQueryToDataTable(sbSQL1.ToString)
        gvBackOrdersSummary.EmptyDataText = "no back orders found"
        gvBackOrdersSummary.DataSource = oDataTable1
        gvBackOrdersSummary.DataBind()
        Dim oDataTable2 As DataTable = ExecuteQueryToDataTable(sSQL2)
        gvBackOrdersDetail.EmptyDataText = "no back orders found"
        gvBackOrdersDetail.DataSource = oDataTable2
        gvBackOrdersDetail.DataBind()
        Call InitddlAddBackOrderProduct()
        ddlAddBackOrderOrganisations.Items.Clear()
        tbAddBackOrderQty.Text = String.Empty
        tbAddBackOrderQty.Enabled = False
    End Sub

    Protected Sub InitddlAddBackOrderProduct()
        Dim sSQL As String = "SELECT ProductCode + ' ' + ProductDescription 'Product', LogisticProductKey FROM LogisticProduct WHERE CustomerKey = 580 AND ArchiveFlag = 'N' AND DeletedFlag = 'N' AND ProductDate = 'generic' ORDER BY ProductCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "Product", "LogisticProductKey")
        ddlAddBackOrderProduct.Items.Clear()
        ddlAddBackOrderProduct.Items.Add(New ListItem("- please select - ", 0))
        For Each li As ListItem In oListItemCollection
            ddlAddBackOrderProduct.Items.Add(li)
        Next
    End Sub
    
    Protected Sub DisplayGoodsIn()
        rbGoodsInMostRecent.Checked = True
        rbGoodsInProductsWithBackOrders.Checked = True
    End Sub

    Protected Sub DisplayAvailableForDistribution()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("CREATE TABLE #temp (LogisticProductKey int, Product varchar(50), Description varchar(200), MinimumStockLevel int, Qty int) ")
        sbSQL.Append("INSERT INTO #temp ")
        sbSQL.Append("SELECT DISTINCT lp.LogisticProductKey, ProductCode, ProductDescription, MinimumStockLevel, ")
        sbSQL.Append("Quantity =     CASE ISNUMERIC((select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("WHEN 0 THEN 0 ELSE (select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("END ")
        sbSQL.Append("FROM LogisticProduct lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        sbSQL.Append("WHERE lp.LogisticProductKey IN (SELECT DISTINCT GenericProductKey FROM NHSPIPLinkedProducts WHERE BacklogQty > 0) ")
        sbSQL.Append("SELECT LogisticProductKey, Product, Description, MinimumStockLevel , Qty FROM #temp WHERE Qty > MinimumStockLevel ")
        sbSQL.Append("DROP TABLE #temp ")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvAvailableForDistribution.DataSource = oDataTable
        gvAvailableForDistribution.DataBind()
    End Sub

    Protected Sub InitStockReservationOrganisations()
        Dim sSQL As String
        If ddlStockReservationsOrganisation.Items.Count > 0 Then
            For i As Integer = ddlStockReservationsOrganisation.Items.Count - 1 To 0 Step -1
                ddlStockReservationsOrganisation.Items.RemoveAt(i)
            Next
        End If
        sSQL = "SELECT PCTName + ' (' + PCTAbbreviation + ')' 'PCTName', [id] FROM NHSPCTs WHERE IsDeleted = 0 AND IsVirtual = 0 AND [id] IN (SELECT DISTINCT NHSPCTKey FROM NHSPIPWebformSubmission) ORDER BY PCTName"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "PCTName", "id")
        ddlStockReservationsOrganisation.Items.Add(New ListItem(" - please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlStockReservationsOrganisation.Items.Add(li)
        Next
    End Sub

    Protected Sub InitActivityLogProductDropdown()
        If ddlActivityLogFilterProduct.Items.Count > 0 Then
            For i As Integer = ddlActivityLogFilterProduct.Items.Count - 1 To 0 Step -1
                ddlActivityLogFilterProduct.Items.RemoveAt(i)
            Next
        End If
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT ProductCode + ' ~ ' + ProductDescription 'Product', LogisticProductKey FROM LogisticProduct WHERE LogisticProductKey IN (SELECT DISTINCT LogisticProductKey FROM NHSPIPActivityLog npal WHERE LogEntry LIKE 'Cloning%') ORDER BY ProductCode", "Product", "LogisticProductKey")
        ddlActivityLogFilterProduct.Items.Add(New ListItem(" - please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlActivityLogFilterProduct.Items.Add(li)
        Next
    End Sub
    
    Protected Sub DisplayActivity()
        Call SetActivityLogControlEnable(False)
        Call ClearActivityLogControls()
        Call InitActivityLogProductDropdown()
    End Sub
    
    Protected Sub ClearActivityLogControls()
        If ddlActivityLogFilterProduct.Items.Count > 0 Then
            ddlActivityLogFilterProduct.SelectedIndex = 0
        End If
        ddlActivityLogFilterProduct.Enabled = False
        tbActivityLogOrderNo.Enabled = False
        tbActivityLogOrderNo.Text = String.Empty
        tbActivityLogText.Enabled = False
        tbActivityLogText.Text = String.Empty
        'rbActivityLogNoFiltering.Checked = True
    End Sub

    Protected Sub DisplayGeneric()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT ProductCode 'Product Code', ProductDescription 'Description', 'Quantity In Stock' = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("END, ")
        sbSQL.Append("upp.MaxGrabQty 'Max Order', LanguageID 'Language' ")
        sbSQL.Append("FROM LogisticProduct AS lp ")
        sbSQL.Append("INNER JOIN UserProductProfile upp ")
        sbSQL.Append("ON lp.LogisticProductKey = upp.ProductKey ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        sbSQL.Append("WHERE CustomerKey = " & gnCustomerKey & " AND ")
        sbSQL.Append("ArchiveFlag = 'N' AND DeletedFlag = 'N' AND ")
        sbSQL.Append("lp.ProductDate = 'GENERIC' AND ")
        sbSQL.Append("upp.UserKey = (SELECT [key] FROM UserProfile WHERE CustomerKey = 580 AND UserId LIKE 'NHSPIP%web%')")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub
    
    Protected Sub DisplayAssignedByOrganisation()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT ProductDate 'Organisation', ProductCode 'Product Code', ProductDescription 'Description', 'Quantity In Stock' = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("END, ")
        sbSQL.Append("LanguageID 'Language' ")
        sbSQL.Append("FROM LogisticProduct AS lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        'sbSQL1.Append("WHERE(lp.LogisticProductKey = ")
        sbSQL.Append("WHERE CustomerKey = " & gnCustomerKey & " AND ")
        sbSQL.Append("ProductDate IN ")
        sbSQL.Append("(SELECT DISTINCT ProductDate FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductDate <> 'GENERIC' AND ISNULL(ProductDate,'') <> '' AND DeletedFlag = 'N') ")
        sbSQL.Append("ORDER BY Organisation ")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub
    
    Protected Sub DisplayAssignedbyCreationDate()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT CONVERT(varchar(9),CreatedOn,6) 'Created', ProductDate 'Organisation', ProductCode 'Product Code', ProductDescription 'Description', 'Quantity In Stock' = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("END, ")
        sbSQL.Append("LanguageID 'Language' ")
        sbSQL.Append("FROM LogisticProduct AS lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        'sbSQL1.Append("WHERE(lp.LogisticProductKey = ")
        sbSQL.Append("WHERE CustomerKey = " & gnCustomerKey & " AND ")
        sbSQL.Append("ProductDate IN ")
        sbSQL.Append("(SELECT DISTINCT ProductDate FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductDate <> 'GENERIC' AND ISNULL(ProductDate,'') <> '' AND DeletedFlag = 'N') ")
        sbSQL.Append("ORDER BY CreatedOn DESC ")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub

    Protected Sub DisplayAwaitingAllocation()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT ProductDate 'Organisation', ProductCode 'Product Code', ProductDescription 'Description', 'Quantity In Stock' = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("END, ")
        sbSQL.Append("LanguageID 'Language' ")
        sbSQL.Append("INTO #temp ")
        sbSQL.Append("FROM LogisticProduct AS lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        'sbSQL1.Append("WHERE(lp.LogisticProductKey = ")
        sbSQL.Append("WHERE CustomerKey = " & gnCustomerKey & " AND ")
        sbSQL.Append("ProductDate IN ")
        sbSQL.Append("(SELECT DISTINCT ProductDate FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND ProductDate <> 'GENERIC' AND ISNULL(ProductDate,'') <> '' AND DeletedFlag = 'N') ")
        sbSQL.Append("ORDER BY Organisation ")
        sbSQL.Append("SELECT * FROM #temp WHERE CAST([Quantity In Stock] AS int) = 0 ")
        sbSQL.Append("DROP TABLE #temp ")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub

    Protected Sub ProductVisibility()
        If ExecuteQueryToDataTable("SELECT up.[key] 'UserKey', up.UserId 'User ID', up.Department 'User Org', ProductCode 'Product Code', ProductDescription 'Description',  ProductDate 'Product Org', ProductKey 'Product Key' FROM UserProductProfile upp INNER JOIN UserProfile up ON upp.UserKey = up.[key] INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = upp.ProductKey WHERE up.CustomerKey = " & gnCustomerKey & " AND up.Status = 'Active' AND lp.Misc1 <> 'GENERIC' AND up.UserId <> 'marilynnhspip' AND up.UserId <> 'nhspip' AND up.UserId <> '' AND AbleToPick = 1 AND lp.DeletedFlag = 'N' AND up.Department <> lp.Misc1 ORDER BY UserId").Rows.Count > 0 Then
            lblVisibilityMismatches.Visible = True
            btnVisibilityMismatches.Visible = True
        End If

        gsSQL = "SELECT up.UserId 'User ID', up.Department 'User Org', ProductCode 'Product Code', ProductDescription 'Description',  ProductDate 'Product Org', ProductKey FROM UserProductProfile upp INNER JOIN UserProfile up ON upp.UserKey = up.[key] INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = upp.ProductKey WHERE up.CustomerKey = " & gnCustomerKey & " AND up.Status = 'Active' AND lp.Misc1 <> 'GENERIC' AND up.UserId <> 'marilynnhspip' AND up.UserId <> 'nhspip' AND up.UserId <> '' AND AbleToPick = 1 AND lp.DeletedFlag = 'N' ORDER BY UserId"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub
    
    Protected Sub btnVisibilityMismatches_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gsSQL = "SELECT up.[key] 'UserKey', up.UserId 'User ID', up.Department 'User Org', ProductCode 'Product Code', ProductDescription 'Description',  ProductDate 'Product Org', ProductKey 'Product Key' FROM UserProductProfile upp INNER JOIN UserProfile up ON upp.UserKey = up.[key] INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = upp.ProductKey WHERE up.CustomerKey = " & gnCustomerKey & " AND up.Status = 'Active' AND lp.Misc1 <> 'GENERIC' AND up.UserId <> 'marilynnhspip' AND up.UserId <> 'nhspip' AND up.UserId <> '' AND AbleToPick = 1 AND lp.DeletedFlag = 'N' AND up.Department <> lp.Misc1 ORDER BY UserId"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        lblVisibilityMismatches.Visible = False
        btnVisibilityMismatches.Visible = False
        lblLegendMismatches.Visible = True
        psDownloadQuery = gsSQL
    End Sub
    
    Protected Sub SearchOrgs()
        
    End Sub
    
    Protected Sub Accounts()
        gsSQL = "SELECT UserID 'User ID', FirstName 'First name', LastName 'Last name', EmailAddr 'Email addr', Department 'Organisation' FROM UserProfile WHERE UserId <> 'NHSPIPPCTWebform' AND UserId <> 'marilynnhspip' AND UserId <> 'nhspip' AND Type = 'User' AND CustomerKey = " & gnCustomerKey & " AND Status = 'Active' ORDER BY Department"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub
    
    Protected Sub ConsistencyCheck()
        gvData.EmptyDataText = "No cross linking discrepancies found in User Accounts list"
        gsSQL = "SELECT [key] 'User #', Department 'Org', UserID 'User ID', FirstName 'First name', LastName 'Last name', EmailAddr 'Email addr' FROM UserProfile WHERE UserId <> 'NHSPIPPCTWebform' AND UserId <> 'marilynnhspip' AND UserId <> 'nhspip' AND Type = 'User' AND CustomerKey = " & gnCustomerKey & " AND Status = 'Active' AND Department NOT IN (SELECT PCTAbbreviation FROM NHSPCTs) ORDER BY Department"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        gvData2.EmptyDataText = "No cross linking discrepancies found in Products list"
        gsSQL = "SELECT LogisticProductKey 'Product #', ISNULL(ProductDate,'') 'Org', ProductCode 'Product Code', ProductDescription 'Description',  CONVERT(VARCHAR(9), LastUpdatedOn, 6) 'Last update' FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND DeletedFlag <> 'Y' AND (NOT (ISNULL(ProductDate,'') = 'GENERIC' OR ISNULL(ProductDate,'') IN (SELECT PCTAbbreviation FROM NHSPCTs)) OR ProductDate IS NULL) ORDER BY LastUpdatedOn DESC"
        Dim oDataTable2 As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData2.DataSource = oDataTable2
        gvData2.DataBind()
    End Sub
    
    Protected Sub NoProducts()
        gsSQL = "SELECT PCTName 'Organisation' FROM NHSPCTs WHERE PCTAbbreviation NOT IN (SELECT DISTINCT ProductDate FROM LogisticProduct WHERE CustomerKey = " & gnCustomerKey & " AND DeletedFlag = 'N' AND ISNULL(Misc1,'') <> '') ORDER BY PCTName"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        gvData.Width = New Unit("50%")
    End Sub
    
    Protected Sub NoAccounts()
        gsSQL = "SELECT PCTName 'Organisation' FROM NHSPCTs WHERE PCTAbbreviation NOT IN (SELECT Department FROM UserProfile WHERE CustomerKey = " & gnCustomerKey & " AND DeletedFlag = 0) ORDER BY PCTName"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        gvData.Width = New Unit("50%")
    End Sub
    
    Protected Sub DisplayOrganisations()
        gsSQL = "SELECT PCTName 'Organisation', PCTAbbreviation 'Abbreviation', PCTCode 'Code' FROM NHSPCTs WHERE [IsDeleted] = 0 "
        If rbShowRealOrganisations.Checked Then
            gsSQL += "AND IsVirtual = 0 "
        Else
            gsSQL += "AND IsVirtual = 1 "
        End If
        gsSQL += "AND " & ddlOrganisationFilter.SelectedValue & " "
        gsSQL += "ORDER BY PCTName"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
    End Sub
        
    Protected Sub rbShowRealOrgs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlOrganisationList.Visible = True
        pnlData.Visible = True
        Call DisplayOrganisations()
    End Sub

    Protected Sub rbShowVirtualOrgs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlOrganisationList.Visible = True
        pnlData.Visible = True
        Call DisplayOrganisations()
    End Sub
    
    Protected Sub ExportData()
        Dim oDatatable As DataTable = ExecuteQueryToDataTable(psDownloadQuery)
        Response.Clear()
        Response.ContentType = "Application/x-msexcel"
                
        Dim sResponseValue As String
        sResponseValue = "attachment; filename=" & ControlChars.Quote & "Dashboard Export " & Now.ToString("ddMMMyyyy - hhmmss") & ".csv" & ControlChars.Quote
        Response.AddHeader("Content-Disposition", sResponseValue)
    
        For Each c As DataColumn In oDatatable.Columns
            Response.Write(c.ColumnName)
            Response.Write(",")
        Next
        Response.Write(vbCrLf)
    
        Dim sItem As String
        For Each dr As DataRow In oDatatable.Rows
            For Each dc As DataColumn In oDatatable.Columns
                sItem = (dr(dc.ColumnName).ToString)
                sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                sItem = ControlChars.Quote & sItem & ControlChars.Quote
                Response.Write(sItem)
                Response.Write(",")
            Next
            Response.Write(vbCrLf)
        Next
        Response.End()
    End Sub
    
    Protected Sub lnkbtnDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExportData()
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

    Property psDownloadQuery() As String
        Get
            Dim o As Object = ViewState("PIP_DownloadQuery")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PIP_DownloadQuery") = Value
        End Set
    End Property

    Protected Sub ddlOrganisationFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DisplayOrganisations()
    End Sub
    
    Protected Sub btnSearchGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        Dim sSearchTerm = "'%" & tbSearchOrgs.Text.Replace("'", "''") & "%'"
        sSQL = "SELECT PCTName 'Organisation', PCTAbbreviation 'Abbreviation', PCTCode 'Code', Type 'Org type' FROM NHSPCTs WHERE PCTName LIKE " & sSearchTerm & " OR PCTAbbreviation LIKE " & sSearchTerm & " OR PCTCode LIKE " & sSearchTerm
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        pnlData.Visible = True
    End Sub
    
    Protected Sub cbStockReservationsLimitByDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetStockReservationControlEnable(cb.Checked)
        If cb.Checked Then
            tbStockReservationFrom.Focus()
        Else
            tbStockReservationFrom.Text = String.Empty
            tbStockReservationTo.Text = String.Empty
        End If
    End Sub
    
    Protected Sub SetStockReservationControlEnable(ByVal bEnabled As Boolean)
        tbStockReservationFrom.Enabled = bEnabled
        tbStockReservationTo.Enabled = bEnabled
        'lnkbtnStockReservationDateLast7Days.Enabled = bEnabled
        'lnkbtnStockReservationDateLast30Days.Enabled = bEnabled
    End Sub

    Protected Sub lnkbtnStockReservationDateLast7Days_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbStockReservationsLimitByDate.Checked = True
        Call SetActivityLogControlEnable(True)
        Call SetStockReservationControlEnable(True)
        tbStockReservationFrom.Text = Today.AddDays(-7).ToString("dd-MMM-yy")
        tbStockReservationTo.Text = Today.AddDays(1).ToString("dd-MMM-yy")
        btnStockReservationsGo.Focus()
    End Sub

    Protected Sub lnkbtnStockReservationDateLast30Days_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbStockReservationsLimitByDate.Checked = True
        Call SetActivityLogControlEnable(True)
        Call SetStockReservationControlEnable(True)
        tbStockReservationFrom.Text = Today.AddDays(-30).ToString("dd-MMM-yy")
        tbStockReservationTo.Text = Today.AddDays(1).ToString("dd-MMM-yy")
        btnStockReservationsGo.Focus()
    End Sub

    Protected Sub rbStockReservationsFilterByOrganisation_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        ddlStockReservationsOrganisation.Enabled = rb.Enabled
    End Sub

    Protected Sub rbStockReservationsNoFiltering_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            ClearReservationsControls()
        End If
    End Sub
    
    Protected Sub ClearReservationsControls()
        If ddlStockReservationsOrganisation.Items.Count > 0 Then
            ddlStockReservationsOrganisation.SelectedIndex = 0
        End If
        ddlStockReservationsOrganisation.Enabled = False
    End Sub

    Protected Sub cbActivityLogLimitByDate_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetActivityLogControlEnable(cb.Checked)
        If cb.Checked Then
            tbActivityLogFrom.Focus()
        Else
            tbActivityLogFrom.Text = String.Empty
            tbActivityLogTo.Text = String.Empty
        End If
    End Sub
    
    Protected Sub SetActivityLogControlEnable(ByVal bEnabled As Boolean)
        tbActivityLogFrom.Enabled = bEnabled
        tbActivityLogTo.Enabled = bEnabled
        'lnkbtnActivityLogDateLast7Days.Enabled = bEnabled
        'lnkbtnActivityLogDateLast30Days.Enabled = bEnabled
    End Sub


    Protected Sub rbActivityLogFilterByStockAdditionFailures_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call ClearActivityLogControls()
        End If
    End Sub

    Protected Sub rbActivityLogFilterByOrderNo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call ClearActivityLogControls()
            tbActivityLogOrderNo.Enabled = True
            tbActivityLogOrderNo.Focus()
        End If
    End Sub

    Protected Sub rbActivityLogFilterByProduct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call ClearActivityLogControls()
            ddlActivityLogFilterProduct.Enabled = True
            ddlActivityLogFilterProduct.Focus()
        End If
    End Sub

    Protected Sub rbActivityLogNoFiltering_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call ClearActivityLogControls()
        End If
    End Sub

    Protected Sub lnkbtnActivityLogDateLast7Days_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbActivityLogLimitByDate.Checked = True
        Call SetActivityLogControlEnable(True)
        tbActivityLogFrom.Text = Today.AddDays(-7).ToString("dd-MMM-yy")
        tbActivityLogTo.Text = Today.AddDays(1).ToString("dd-MMM-yy")
        btnActivityLogGo.Focus()
    End Sub

    Protected Sub lnkbtnActivityLogDateLast30Days_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        cbActivityLogLimitByDate.Checked = True
        Call SetActivityLogControlEnable(True)
        tbActivityLogFrom.Text = Today.AddDays(-30).ToString("dd-MMM-yy")
        tbActivityLogTo.Text = Today.AddDays(1).ToString("dd-MMM-yy")
        btnActivityLogGo.Focus()
    End Sub
    
    Protected Function CheckDateRange(ByVal sFromDate As String, ByVal sToDate As String) As String
        CheckDateRange = String.Empty
        Dim dtFromDate As Date
        Dim dtToDate As Date
        If Not IsDate(sFromDate) Then
            CheckDateRange = "FROM date is not a valid date"
            Exit Function
        Else
            dtFromDate = Date.Parse(sFromDate)
        End If
        If Not IsDate(sToDate) Then
            CheckDateRange = "TO date is not a valid date"
            Exit Function
        Else
            dtToDate = Date.Parse(sToDate)
        End If
        If dtFromDate >= dtToDate Then
            CheckDateRange = "FROM date must be earlier than TO date"
            Exit Function
        End If
        If dtFromDate > Date.Today Then
            CheckDateRange = "FROM date is beyond the current date"
            Exit Function
        End If
    End Function
    
    Protected Sub btnStockReservationsGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call StockReservationsGo()
    End Sub
    
    Protected Sub StockReservationsGo()
        If rbStockReservationsFilterByOrganisation.Checked AndAlso ddlStockReservationsOrganisation.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select the organisation to filter on.")
            Exit Sub
        End If
        If cbStockReservationsLimitByDate.Checked Then
            Dim sCheckDateMessage As String = CheckDateRange(tbStockReservationFrom.Text, tbStockReservationTo.Text)
            If sCheckDateMessage <> String.Empty Then
                WebMsgBox.Show(sCheckDateMessage)
                Exit Sub
            End If
        End If
        If Not (cbStockReservationsFullyAssigned.Checked Or cbStockReservationsPartiallyAssigned.Checked Or cbStockReservationsUnassigned.Checked) Then
            WebMsgBox.Show("You have not selected any reservations to show (fully/partially/unassigned)")
            Exit Sub
        End If
        gsSQL = "SELECT CONVERT(VARCHAR(9), nws.CreatedOn, 6) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), nws.CreatedOn, 108)),1,5) 'Created on', nws.[id] 'Order #', org.PCTName + ' (' + org.PCTAbbreviation + ')' 'Organisation', lp.ProductCode 'Product code', lp.ProductDate 'Product date', lp.ProductDescription 'Description', nwsd.QtyRequested 'Qty requested', nwsd.QtyAvailable 'Qty available', FirstName 'First name', LastName 'Last name', EmailAddr 'Email addr', AccountName 'Account' FROM NHSPIPWebformSubmission nws INNER JOIN NHSPIPWebformSubmissionDetail nwsd ON nws.[id] = nwsd.OrderKey INNER JOIN LogisticProduct lp ON lp.LogisticProductKey = nwsd.LogisticProductKey INNER JOIN NHSPCTs org ON nws.NHSPCTKey = org.[id] WHERE 1 = 1 "
        If cbStockReservationsLimitByDate.Checked Then
            gsSQL += " AND nws.CreatedOn >= '" & tbStockReservationFrom.Text & "' AND nws.CreatedOn <= '" & tbStockReservationTo.Text & "' "
        End If
        If rbStockReservationsFilterByOrganisation.Checked Then
            gsSQL += " AND NHSPCTkey = " & ddlStockReservationsOrganisation.SelectedValue & " "
        End If
        Dim sSQL2 As String = String.Empty
        If Not (cbStockReservationsFullyAssigned.Checked And cbStockReservationsPartiallyAssigned.Checked And cbStockReservationsUnassigned.Checked) Then
            If cbStockReservationsFullyAssigned.Checked And cbStockReservationsPartiallyAssigned.Checked Then
                sSQL2 = "AND nwsd.QtyAvailable > 0 "
            ElseIf cbStockReservationsFullyAssigned.Checked And cbStockReservationsUnassigned.Checked Then
                sSQL2 = "AND (nwsd.QtyAvailable >= nwsd.QtyRequested OR nwsd.QtyAvailable = 0) "
            ElseIf cbStockReservationsPartiallyAssigned.Checked And cbStockReservationsUnassigned.Checked Then
                sSQL2 = "AND nwsd.QtyAvailable < nwsd.QtyRequested "
            ElseIf cbStockReservationsFullyAssigned.Checked Then
                sSQL2 = "AND nwsd.QtyRequested <= nwsd.QtyAvailable "
            ElseIf cbStockReservationsPartiallyAssigned.Checked Then
                sSQL2 = "AND nwsd.QtyRequested < nwsd.QtyAvailable AND nwsd.QtyAvailable > 0 "
            ElseIf cbStockReservationsUnassigned.Checked Then
                sSQL2 = "AND nwsd.QtyAvailable = 0 "
            End If
        End If
        gsSQL += sSQL2 & " ORDER BY nws.[id]"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        pnlData.Visible = True
    End Sub

    Protected Sub btnActivityLogGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ActivityLogGo()
    End Sub
    
    Protected Sub ActivityLogGo()
        tbActivityLogOrderNo.Text = tbActivityLogOrderNo.Text.Trim
        tbActivityLogText.Text = tbActivityLogText.Text.Trim
        If rbActivityLogFilterByOrderNo.Checked AndAlso (tbActivityLogOrderNo.Text = String.Empty OrElse Not IsNumeric(tbActivityLogOrderNo.Text)) Then
            WebMsgBox.Show("Please specify the order number to filter on.")
            Exit Sub
        End If
        If rbActivityLogFilterByProduct.Checked AndAlso ddlActivityLogFilterProduct.SelectedIndex = 0 Then
            WebMsgBox.Show("Please specify the product to filter on.")
            Exit Sub
        End If
        If cbActivityLogLimitByDate.Checked Then
            Dim sCheckDateMessage As String = CheckDateRange(tbActivityLogFrom.Text, tbActivityLogTo.Text)
            If sCheckDateMessage <> String.Empty Then
                WebMsgBox.Show(sCheckDateMessage)
                Exit Sub
            End If
        End If
        If rbActivityLogActivityLogText.Checked AndAlso tbActivityLogText.Text = String.Empty Then
            WebMsgBox.Show("Please specify the log message text to filter on.")
            Exit Sub
        End If
        gsSQL = "SELECT * FROM NHSPIPActivityLog npal WHERE 1 = 1 "
        If cbActivityLogLimitByDate.Checked Then
            gsSQL += " AND npal.CreatedOn >= '" & tbActivityLogFrom.Text & "' AND npal.CreatedOn <= '" & tbActivityLogTo.Text & "' "
        End If
        Dim sSQL2 As String = String.Empty
        If rbActivityLogFilterByStockAdditionFailures.Checked Then
            sSQL2 = "AND (npal.LogEntry LIKE '%insufficient pick quantity%' OR  npal.LogEntry LIKE '%no pick quantity%')"
        ElseIf rbActivityLogFilterByOrderNo.Checked Then
            sSQL2 = "AND npal.OrderNo = " & tbActivityLogOrderNo.Text & " "
        ElseIf rbActivityLogFilterByProduct.Checked Then
            sSQL2 = "AND (npal.LogisticProductKey = " & ddlActivityLogFilterProduct.SelectedValue & " OR npal.LogisticProductKey IN (SELECT RingFencedProductKey FROM NHSPIPLinkedProducts WHERE GenericProductKey = " & ddlActivityLogFilterProduct.SelectedValue & ")) "
        ElseIf rbActivityLogActivityLogText.Checked Then
            sSQL2 = "AND npal.LogEntry LIKE '%" & tbActivityLogText.Text.Replace("'", "''") & "%' "
        End If
        gsSQL += sSQL2 & "ORDER BY [id]"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        pnlData.Visible = True
    End Sub

    Protected Sub rbActivityLogActivityLogText_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            Call ClearActivityLogControls()
            tbActivityLogText.Enabled = True
            tbActivityLogText.Focus()
        End If
    End Sub
    
    Protected Sub btnClearAllOrganisationProductBackOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nProductKey As Integer = b.CommandArgument
        Dim sSQL As String = "UPDATE NHSPIPLinkedProducts SET BacklogQty = 0 WHERE GenericProductKey = " & nProductKey & " DELETE FROM NHSPIPBackOrderNote WHERE LogisticProductKey = " & nProductKey
        Call ExecuteNonQuery(sSQL)
        Call Log(LOG_ENTRY_TYPE_PRODUCT, nProductKey, -1, String.Empty, "Cleared master product back order")
        Call DisplayBackOrders()
    End Sub

    Protected Sub btnClearOrganisationProductBackOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLinkedProductKey As Integer = b.CommandArgument
        Dim sSQL As String = "UPDATE NHSPIPLinkedProducts SET BacklogQty = 0 WHERE RingFencedProductKey = " & nLinkedProductKey & " DELETE FROM NHSPIPBackOrderNote WHERE LogisticProductKey = " & nLinkedProductKey
        Call ExecuteNonQuery(sSQL)
        Call Log(LOG_ENTRY_TYPE_PRODUCT, nLinkedProductKey, -1, String.Empty, "Cleared organisation product back order")
        Call DisplayBackOrders()
    End Sub
    
    Protected Sub lnkbtnAddBackOrderNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Dim oDataTable As DataTable
        Dim nLogisticProductKey As Integer = lb.CommandArgument
        Call HideAllPanels()
        pnlAddNote.Visible = True
        oDataTable = ExecuteQueryToDataTable("SELECT ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey)
        Dim sProductCode As String = oDataTable.Rows(0).Item("ProductCode")
        Dim sProductDate As String = oDataTable.Rows(0).Item("ProductDate")
        If sProductDate.ToLower.Contains("generic") Then
            lblProductNote.Text = " MASTER product " & sProductCode
        Else
            lblProductNote.Text = " product " & sProductCode & " (" & sProductDate & ")"
        End If
        lblProductNote.Text += " - " & oDataTable.Rows(0).Item("ProductDescription")
        oDataTable = ExecuteQueryToDataTable("SELECT NoteText FROM NHSPIPBackOrderNote WHERE LogisticProductKey = " & nLogisticProductKey)
        If oDataTable.Rows.Count > 0 Then
            lblProductCurrentNote.Text = oDataTable.Rows(0).Item(0)
        Else
            lblProductCurrentNote.Text = ""
        End If
        btnSaveNote.CommandArgument = nLogisticProductKey
        tbNote.Text = String.Empty
        tbNote.Focus()
    End Sub

    Protected Sub btnAddBackOrderNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim oDataTable As DataTable
        Dim nLogisticProductKey As Integer = b.CommandArgument
        Call HideAllPanels()
        pnlAddNote.Visible = True
        oDataTable = ExecuteQueryToDataTable("SELECT ProductCode, ProductDate, ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & nLogisticProductKey)
        Dim sProductCode As String = oDataTable.Rows(0).Item("ProductCode")
        Dim sProductDate As String = oDataTable.Rows(0).Item("ProductDate")
        If sProductDate.ToLower.Contains("generic") Then
            lblProductNote.Text = " MASTER product " & sProductCode
        Else
            lblProductNote.Text = " product " & sProductCode & " (" & sProductDate & ")"
        End If
        lblProductNote.Text += " - " & oDataTable.Rows(0).Item("ProductDescription")
        oDataTable = ExecuteQueryToDataTable("SELECT NoteText FROM NHSPIPBackOrderNote WHERE LogisticProductKey = " & nLogisticProductKey)
        If oDataTable.Rows.Count > 0 Then
            lblProductCurrentNote.Text = oDataTable.Rows(0).Item(0)
        Else
            lblProductCurrentNote.Text = ""
        End If
        btnSaveNote.CommandArgument = nLogisticProductKey
        tbNote.Text = String.Empty
        tbNote.Focus()
    End Sub

    Protected Sub btnSaveNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sSQL As String
        Dim nLogisticProductKey As Integer = b.CommandArgument
        pnlAddNote.Visible = False
        Dim sNoteText As String = lblProductCurrentNote.Text
        Dim sDate As String = "<b>" & Date.Now.ToString("ddMMMyy hh:mm") & "</b> "
        sNoteText += sDate & tbNote.Text & "<br />"
        If lblProductCurrentNote.Text <> String.Empty Then
            sSQL = "UPDATE NHSPIPBackOrderNote SET NoteText = '" & sNoteText.Replace("'", "''") & "' WHERE LogisticProductKey = " & nLogisticProductKey.ToString
        Else
            sSQL = "INSERT INTO NHSPIPBackOrderNote (LogisticProductKey, NoteText) VALUES (" & nLogisticProductKey.ToString & ", '" & sNoteText.Replace("'", "''") & "')"
        End If
        Call ExecuteNonQuery(sSQL)
        pnlBackOrders.Visible = True
        pnlBackOrdersData.Visible = True
        Call Log(LOG_ENTRY_TYPE_PRODUCT, nLogisticProductKey, -1, String.Empty, "Added note to product " & nLogisticProductKey.ToString & ": " & tbNote.Text)
        Call DisplayBackOrders()
    End Sub

    Protected Sub btnCancelNote_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call RefreshBackOrders()
    End Sub

    Protected Sub RefreshBackOrders()
        Call HideAllPanels()
        pnlBackOrders.Visible = True
        pnlBackOrdersData.Visible = True
        Call DisplayBackOrders()
    End Sub
    
    Protected Sub Log(ByVal sType As String, ByVal nLogisticProductKey As Integer, ByVal nOrderNo As Integer, ByVal sPCTAbbreviation As String, ByVal sLogEntry As String)
        Dim sbSQL As New StringBuilder
        sbSQL.Append("INSERT INTO NHSPIPActivityLog (Type, LogisticProductKey, OrderNo, PCTAbbreviation, LogEntry, CreatedOn) VALUES (")
        
        sbSQL.Append("'")
        sbSQL.Append(sType)
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append(nLogisticProductKey)
        sbSQL.Append(",")

        sbSQL.Append(nOrderNo)
        sbSQL.Append(",")

        sbSQL.Append("'")
        sbSQL.Append(sPCTAbbreviation.Replace("'", "''"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("'")
        sbSQL.Append(sLogEntry.Replace("''", "'"))
        sbSQL.Append("'")
        sbSQL.Append(",")
        
        sbSQL.Append("GETDATE()")
        sbSQL.Append(")")
        Call ExecuteNonQuery(sbSQL.ToString)
    End Sub
    
    Protected Sub btnGoodsIn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GoodsIn()
    End Sub
    
    Protected Sub GoodsIn()
        Dim sAllMasterProducts As String = " SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = 580 AND ProductDate = 'generic' AND DeletedFlag = 'N' and ArchiveFlag = 'N' "
        Dim sMasterProductsWithBacklog As String = " SELECT DISTINCT GenericProductKey FROM NHSPIPLinkedProducts WHERE BacklogQty > 0 "
        Dim sbSQL As New StringBuilder
        
        sbSQL.Append("CREATE TABLE #temp (Date smalldatetime, Product varchar(50), Description varchar(200), Qty int) ")
        sbSQL.Append("DECLARE @x int ")
        sbSQL.Append("DECLARE c CURSOR FOR ")
        If rbGoodsInProductsAll.Checked Then
            sbSQL.Append(sAllMasterProducts)
        Else
            sbSQL.Append(sMasterProductsWithBacklog)
        End If
        ' sbSQL.Append("  SELECT LogisticProductKey FROM LogisticProduct WHERE CustomerKey = 580 AND ProductDate = 'generic' AND DeletedFlag = 'N' and ArchiveFlag = 'N' ")
        sbSQL.Append("  OPEN c ")
        sbSQL.Append("  FETCH NEXT FROM c INTO @x ")
        sbSQL.Append("  WHILE (@@FETCH_STATUS) = 0 ")
        sbSQL.Append("  BEGIN ")
        sbSQL.Append("    INSERT INTO #temp ")
        sbSQL.Append("    SELECT TOP ")
        If rbGoodsInMostRecent.Checked Then
            sbSQL.Append("1")
        Else
            sbSQL.Append("1000")
        End If
        sbSQL.Append("    LogisticMovementStartDateTime, lp.ProductCode, lp.ProductDescription, ItemsIn FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey WHERE LogisticMovementStateId = 'goods-in' and lm.LogisticProductKey = @x ORDER BY LogisticMovementStartDateTime DESC ")
        sbSQL.Append("    FETCH NEXT FROM c INTO @x ")
        sbSQL.Append("  END ")
        sbSQL.Append("CLOSE c ")
        sbSQL.Append("DEALLOCATE c ")
        sbSQL.Append("SELECT CONVERT(varchar(9),Date,6) 'Goods in date', Product, Description, Qty 'Qty in' FROM #temp ORDER BY Product ")
        sbSQL.Append("DROP TABLE #temp ")
        gsSQL = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvData.DataSource = oDataTable
        gvData.DataBind()
        pnlData.Visible = True
        psDownloadQuery = gsSQL
        pnlDownload.Visible = True
    End Sub
    
    Protected Sub btnDistribute_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLogisticProductKey As Integer = b.CommandArgument
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT RingFencedProductKey, ProductDate, ")
        sbSQL.Append("Quantity = CASE ISNUMERIC((select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL.Append("           WHEN 0 THEN 0 ELSE (select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL.Append("           END, ")
        sbSQL.Append("BacklogQty ")
        sbSQL.Append("FROM NHSPIPLinkedProducts nplp ")
        sbSQL.Append("INNER JOIN LogisticProduct lp ")
        sbSQL.Append("ON nplp.RingFencedProductKey = lp.LogisticProductKey ")
        sbSQL.Append("WHERE BacklogQty > 0 ")
        sbSQL.Append("AND GenericProductKey = ")
        sbSQL.Append(nLogisticProductKey.ToString)
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sbSQL.ToString)
        gvDistribute.DataSource = oDataTable
        gvDistribute.DataBind()
        Dim sbSQL2 As New StringBuilder
        sbSQL2.Append("SELECT ProductCode, ProductDescription, ")
        sbSQL2.Append("Quantity = CASE ISNUMERIC((select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey)) ")
        sbSQL2.Append("           WHEN 0 THEN 0 ELSE (select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey) ")
        sbSQL2.Append("           END ")
        sbSQL2.Append("FROM LogisticProduct lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticProductLocation AS lpl ")
        sbSQL.Append("ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        sbSQL2.Append("WHERE lp.LogisticProductKey = ")
        sbSQL2.Append(nLogisticProductKey)
        Dim oDataTable2 As DataTable
        oDataTable2 = ExecuteQueryToDataTable(sbSQL2.ToString)
        Dim dr As DataRow = oDataTable2.Rows(0)
        lblDistributeProduct.Text = dr("ProductCode") & " " & dr("ProductDescription")
        lblDistributeAvailableStockQty.Text = dr("Quantity")
        lblDistributeStockLevelAfterDistribution.Text = String.Empty
        btnStartDistribution.CommandArgument = nLogisticProductKey
        btnStartDistribution.CommandName = dr("ProductCode")
        Call Recalculate()
        Call HideAllPanels()
        pnlDistribute.Visible = True
    End Sub
    
    Protected Function CheckDistribution() As Integer
        CheckDistribution = -1
        lblDistributeStockLevelAfterDistribution.Text = "????"
        Dim nTotalToTransfer As Integer
        Dim tb As TextBox
        For Each gvr As GridViewRow In gvDistribute.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                tb = gvr.FindControl("tbQtyToTransfer")
                If IsNumeric(tb.Text) Then
                    nTotalToTransfer += CInt(tb.Text)
                Else
                    Exit Function
                End If
            End If
        Next
        CheckDistribution = nTotalToTransfer
    End Function
    
    Protected Sub UpdateStockLevelAfterDistribution()
        Dim nStockLevelAfterDistribution As Integer = CInt(lblDistributeAvailableStockQty.Text) - CheckDistribution()
        If nStockLevelAfterDistribution > 0 Then
            lblDistributeStockLevelAfterDistribution.Text = nStockLevelAfterDistribution.ToString
        End If
    End Sub
    
    Protected Sub lnkbtnTransfer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nAmountTaggedForTransfer As Integer = CheckDistribution()

        If nAmountTaggedForTransfer < 0 Then
            WebMsgBox.Show("Format error in one or more quantities to transfer")
            Exit Sub
        End If

        If nAmountTaggedForTransfer > CInt(lblDistributeAvailableStockQty.Text) Then
            WebMsgBox.Show("More stock is marked for transfer than is available for transfer!!")
            Exit Sub
        End If
        
        Dim lb As LinkButton = sender
        Dim nIndex As Integer = lb.CommandArgument
        Dim lbl As Label
        Dim tb As TextBox
        Dim cb As CheckBox
        Dim gvr As GridViewRow = gvDistribute.Rows(nIndex)
        lbl = gvr.FindControl("lblBackLogQty")
        tb = gvr.FindControl("tbQtyToTransfer")
        If (nAmountTaggedForTransfer + CInt(lbl.Text)) <= CInt(lblDistributeAvailableStockQty.Text) Then
            tb.Text = lbl.Text
            cb = gvr.FindControl("cbCloseBackOrder")
            cb.Checked = True
        Else
            WebMsgBox.Show("Transferring this amount would exceed the amount of available MASTER product stock")
        End If
        Call UpdateStockLevelAfterDistribution()
    End Sub
    
    Protected Sub gvDistribute_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            ' Dim tc As TableCell = gvr.Cells(2)
            Dim lbl As Label
            Dim lb As LinkButton
            'lbl = tc.FindControl("lblBackLogQty")
            'lb = tc.FindControl("lnkbtnTransfer")
            lbl = gvr.FindControl("lblBackLogQty")
            lb = gvr.FindControl("lnkbtnTransfer")
            lb.Text = "transfer " & lbl.Text & "-->"
        End If
    End Sub

    Protected Function GetWebformUserKey() As Integer
        Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT [key] FROM UserProfile WHERE UserId = '" & gsWebFormControlUserId & "'")
        GetWebformUserKey = oDataTable.Rows(0).Item(0)
    End Function
    
    Protected Function GeneratePick(ByVal nLogisticProductKey As Integer, ByVal nQty As Integer, ByVal sDestinationProduct As String) As String   ' returns consignment key (numeric, = success) or message (error, failure)
        GeneratePick = String.Empty
        Dim sSQL As String
        Dim oDataTable As DataTable
        sSQL = "SELECT ISNULL(CustomerName,''), ISNULL(CustomerAddr1,''), ISNULL(CustomerAddr2,''), ISNULL(CustomerAddr3,''), ISNULL(CustomerTown,''), ISNULL(CustomerCounty,''), ISNULL(CustomerPostCode,''), ISNULL(CustomerCountryKey,0) FROM Customer WHERE CustomerKey = 5"
        oDataTable = ExecuteQueryToDataTable(sSQL)
        Dim oDataRow As DataRow = oDataTable.Rows(0)
        
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure

        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = GetWebformUserKey()
        oCmdAddBooking.Parameters.Add(param1)
        
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = gnCustomerKey
        oCmdAddBooking.Parameters.Add(param2)

        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING"
        oCmdAddBooking.Parameters.Add(param2a)

        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        param3.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param3)

        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        param4.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param4)

        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        param5.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param5)

        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)
        param6.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param6)
            
        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
            
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = "AUTOMATIC PICK: Please transfer picked items to NHS PIP product " & sDestinationProduct
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param8)

        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = "STOCK ITEM"
        oCmdAddBooking.Parameters.Add(param9)

        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)

        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = "INTERNAL TRANSFER"
        oCmdAddBooking.Parameters.Add(param11)

        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = oDataRow(0)
        oCmdAddBooking.Parameters.Add(param13)
        
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = oDataRow(1)
        oCmdAddBooking.Parameters.Add(param14)
        
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = oDataRow(2)
        
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = oDataRow(3)
        
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = oDataRow(4)
        
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = oDataRow(5)
        
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = oDataRow(6)
        
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = oDataRow(7)
        
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = "Marilyn Quinn X506"
        oCmdAddBooking.Parameters.Add(param21)
        
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = "020 8751 1111"
        oCmdAddBooking.Parameters.Add(param22)

        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = "m.quinn@sprintexpress.co.uk"
        oCmdAddBooking.Parameters.Add(param23)
        
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = oDataRow(0)
        oCmdAddBooking.Parameters.Add(param25)
        
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = oDataRow(1)
        oCmdAddBooking.Parameters.Add(param26)
        
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = oDataRow(2)
        oCmdAddBooking.Parameters.Add(param27)
        
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = oDataRow(3)
        oCmdAddBooking.Parameters.Add(param28)
        
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = oDataRow(4)
        oCmdAddBooking.Parameters.Add(param29)
        
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = oDataRow(5)
        oCmdAddBooking.Parameters.Add(param30)
        
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = oDataRow(6)
        oCmdAddBooking.Parameters.Add(param31)
        
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = oDataRow(7)
        oCmdAddBooking.Parameters.Add(param32)
        
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = "Marilyn Quinn X506"
        oCmdAddBooking.Parameters.Add(param33)
        
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = String.Empty
        oCmdAddBooking.Parameters.Add(param34)
        
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = String.Empty
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
                Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                
                Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                param51.Value = CLng(GetWebformUserKey())
                oCmdAddStockItem.Parameters.Add(param51)
                
                Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                param52.Value = CLng(gnCustomerKey)
                oCmdAddStockItem.Parameters.Add(param52)
                
                Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                param53.Value = lBookingKey
                oCmdAddStockItem.Parameters.Add(param53)
                
                Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                param54.Value = nLogisticProductKey
                oCmdAddStockItem.Parameters.Add(param54)
                
                Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                param55.Value = "PENDING"
                oCmdAddStockItem.Parameters.Add(param55)
                
                Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                param56.Value = nQty
                oCmdAddStockItem.Parameters.Add(param56)
                
                Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                param57.Value = lConsignmentKey
                oCmdAddStockItem.Parameters.Add(param57)
                
                oCmdAddStockItem.Connection = oConn
                oCmdAddStockItem.Transaction = oTrans
                oCmdAddStockItem.ExecuteNonQuery()

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
                GeneratePick = "Zero booking key returned"
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                GeneratePick = lConsignmentKey.ToString
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            GeneratePick = "-> " & ex.ToString
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Function

    Protected Sub lnkbtnRecalculate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Recalculate()
    End Sub
    
    Protected Sub Recalculate()
        If CheckDistribution() >= 0 Then
            Call UpdateStockLevelAfterDistribution()
        End If
    End Sub
    
    Protected Sub ddlAddBackOrderProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue > 0 Then
            Dim sProductCode As String = ExecuteQueryToDataTable("SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & ddlAddBackOrderProduct.SelectedValue).Rows(0).Item(0)
            Call InitddlAddBackOrderOrganisations(sProductCode)
        Else
            ddlAddBackOrderOrganisations.Items.Clear()
            tbAddBackOrderQty.Text = String.Empty
        End If
    End Sub

    Protected Sub InitddlAddBackOrderOrganisations(ByVal sProduct As String)
        Dim sSQL As String = "SELECT ProductDate FROM LogisticProduct WHERE CustomerKey = 580 AND ProductDate <> 'generic' AND ArchiveFlag = 'N' AND DeletedFlag = 'N' AND ProductCode LIKE '" & sProduct.Replace("'", "''") & "'"
        Dim oListItemCollection As ListItemCollection
        oListItemCollection = ExecuteQueryToListItemCollection(sSQL, "ProductDate", "ProductDate")
        ddlAddBackOrderOrganisations.Items.Clear()
        ddlAddBackOrderOrganisations.Items.Add(New ListItem("- please select - ", 0))
        For Each li As ListItem In oListItemCollection
            ddlAddBackOrderOrganisations.Items.Add(li)
        Next
    End Sub
    
    Protected Function GetProductKeyFromProductCodeAndProductDate(ByVal sProductCode As String, ByVal sProductDate As String) As Integer
        GetProductKeyFromProductCodeAndProductDate = 0
        Dim sSQL As String
        sSQL = "SELECT LogisticProductKey FROM LogisticProduct WHERE ProductCode = '" & sProductCode.Replace("'", "''") & "' AND ProductDate = '" & sProductDate.Replace("'", "''") & "'"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count = 1 Then
            GetProductKeyFromProductCodeAndProductDate = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        End If
        'sSQL = "SELECT LogisticProductKey FROM LogisticProduct WHERE ProductCode = '" & sProductCode.Replace("'", "''") & "' AND ProductDate = '" & ddlAddBackOrderOrganisations.SelectedValue.Replace("'", "''") & "'"
        'Dim nRingFencedProductKey As Integer = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
    End Function
    
    Protected Sub InitAddBackOrderQty()
        Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & ddlAddBackOrderProduct.SelectedValue
        Dim sProductCode As String = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        Dim nRingFencedProductKey As Integer = GetProductKeyFromProductCodeAndProductDate(sProductCode, ddlAddBackOrderOrganisations.SelectedValue)
        sSQL = "SELECT BacklogQty FROM NHSPIPLinkedProducts WHERE GenericProductKey = " & ddlAddBackOrderProduct.SelectedValue & " AND RingFencedProductKey = " & nRingFencedProductKey
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count = 1 Then
            tbAddBackOrderQty.Text = oDataTable.Rows(0).Item(0)
        Else
            tbAddBackOrderQty.Text = 0
        End If
        btnAddBackOrderGo.CommandArgument = nRingFencedProductKey
    End Sub
    
    Protected Sub ddlAddBackOrderOrganisations_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        tbAddBackOrderQty.Text = String.Empty
        If ddl.SelectedValue <> "0" Then
            tbAddBackOrderQty.Enabled = True
            tbAddBackOrderQty.Focus()
            Call InitAddBackOrderQty()
        Else
            tbAddBackOrderQty.Enabled = False
        End If
    End Sub
    
    Protected Sub btnAddBackOrderGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbAddBackOrderQty.Text) Then
            WebMsgBox.Show("Quantity to add must be a positive whole number!")
            Exit Sub
        End If
        Dim b As Button = sender
        Dim nRingFencedProductKey As Integer = b.CommandArgument
        Dim sSQL As String
        Dim nGenericProductKey As Integer = ddlAddBackOrderProduct.SelectedValue
        sSQL = "SELECT BacklogQty FROM NHSPIPLinkedProducts WHERE GenericProductKey = " & nGenericProductKey & " AND RingFencedProductKey = " & nRingFencedProductKey
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count = 1 Then
            sSQL = "UPDATE NHSPIPLinkedProducts SET BacklogQty = " & tbAddBackOrderQty.Text & " WHERE GenericProductKey = " & nGenericProductKey & " AND RingFencedProductKey = " & nRingFencedProductKey
        Else
            sSQL = "INSERT INTO NHSPIPLinkedProducts (GenericProductKey, RingFencedProductKey, BacklogQty) VALUES (" & nGenericProductKey & ", " & nRingFencedProductKey & ", " & tbAddBackOrderQty.Text & ")"
        End If
        Call ExecuteNonQuery(sSQL)
        Call RefreshBackOrders()
    End Sub
    
    Protected Sub btnStartDistribution_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim nAmountTaggedForTransfer As Integer = CheckDistribution()
        If nAmountTaggedForTransfer < 0 Then
            WebMsgBox.Show("Format error in one or more quantities to transfer")
            Exit Sub
        End If
        If nAmountTaggedForTransfer > CInt(lblDistributeAvailableStockQty.Text) Then
            WebMsgBox.Show("More stock is marked for transfer than is available for transfer!!")
            Exit Sub
        End If
        If nAmountTaggedForTransfer = 0 Then
            WebMsgBox.Show("No stock has been marked for transfer!")
            Exit Sub
        End If
        Dim nGenericProductKey As Integer = btnStartDistribution.CommandArgument
        Dim sProductName As String = btnStartDistribution.CommandName
        Dim tb As TextBox
        Dim cb As CheckBox
        For Each gvr As GridViewRow In gvDistribute.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                tb = gvr.FindControl("tbQtyToTransfer")
                If IsNumeric(tb.Text) Then
                    If CInt(tb.Text) > 0 Then
                        Call GeneratePick(nGenericProductKey, tb.Text, "Product code: " & sProductName & " Product Date: " & gvr.Cells(0).Text)
                        Call Log(LOG_ENTRY_TYPE_PRODUCT, nGenericProductKey, -1, gvr.Cells(0).Text, "Dashboard generated pick for " & sProductName & ", " & gvr.Cells(0).Text)
                    End If
                End If
                cb = gvr.FindControl("cbCloseBackOrder")
                If cb.Checked Then
                    Dim nRingFencedProductKey = GetProductKeyFromProductCodeAndProductDate(sProductName, gvr.Cells(0).Text)
                    Call ExecuteNonQuery("UPDATE NHSPIPLinkedProducts SET BacklogQty = 0 WHERE GenericProductKey = " & nGenericProductKey & " AND RingFencedProductKey = " & nRingFencedProductKey)
                End If
            End If
        Next
        Call RefreshBackOrders()
    End Sub
    
    Protected Sub rbOrdersAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            If ddlOrdersOriginator.Items.Count > 0 Then
                ddlOrdersOriginator.SelectedIndex = 0
                ddlOrdersOriginator.Enabled = False
            End If
        End If
    End Sub

    Protected Sub rbOrdersPlacedBy_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            ddlOrdersOriginator.Enabled = True
            If ddlOrdersOriginator.Items.Count = 0 Then
                Call InitOrdersOriginatorDropdown()
            End If
        End If
    End Sub

    Protected Sub InitOrdersOriginatorDropdown()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT UserKey, StateId, Description INTO #temp FROM Consignment WHERE CustomerKey = 580 ")
        sbSQL.Append("DELETE FROM #temp WHERE StateId = 'CANCELLED' OR StateId IS NULL ")
        sbSQL.Append("DELETE FROM #temp WHERE Description = 'INTERNAL TRANSFER' ")
        sbSQL.Append("DELETE FROM #temp WHERE UserKey IN (5850, 11932, 11912, 11576) ")
        sbSQL.Append("SELECT DISTINCT dbo.NoOrg(Department) + ' - ' + UserId + ' (' + FirstName + ' ' + LastName + ')' 'User', UserKey FROM #temp t INNER JOIN UserProfile up ON t.userkey = up.[key] ")
        Dim olistItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sbSQL.ToString, "User", "UserKey")
        ddlOrdersOriginator.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In olistItemCollection
            ddlOrdersOriginator.Items.Add(li)
        Next
    End Sub
    
    Protected Sub btnOrderGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        If rbOrdersPlacedBy.Checked Then
            If ddlOrdersOriginator.SelectedIndex = 0 Then
                WebMsgBox.Show("Please select a user.")
                Exit Sub
            End If
        End If
        Dim sbSQL As New StringBuilder
        sbSQL.Append("CREATE TABLE #temp ( ")
        sbSQL.Append("ConsignmentDate smalldatetime, AWB varchar(50), ConsignmentWeight decimal(8,3), ConsignmentAmount money, ConsignmentNOP int, ConsignmentMode varchar(20), ")
        sbSQL.Append("CustomerRef1 nvarchar(30), CustomerRef2 nvarchar(30), Misc1 nvarchar(50), Misc2 nvarchar(50), SpecialInstructions nvarchar(50), ")
        sbSQL.Append("ProductCode varchar(50), ProductDate varchar(50), ProductDescription varchar(200), ItemsOut int, ProductCategory varchar(50), UnitWeightGrams int, ")
        sbSQL.Append("UnitValue money, UserId varchar(250), UserDepartmentCostCentre nvarchar(20), Consignee varchar(1000), CneeTown nvarchar(50), CneeCountry nvarchar(50), ")
        sbSQL.Append("AgentName varchar(50), DespatchDate varchar(50), ExternalSystemId nvarchar(50), ")
        sbSQL.Append("InvoiceNumber varchar(50), InvoiceDate varchar(100), PODInfo varchar(50), PODDate varchar(50) ")
        sbSQL.Append(") ON [PRIMARY] ")
        sbSQL.Append("DECLARE @ConsignmentKey int, @BookedByKey int, @ConsignmentDate smalldatetime, @ConsignmentWeight decimal(8,3), @ConsignmentAmount money, @ConsignmentNOP int, @ConsignmentMode varchar(20), @UserId as varchar(50), @UserDepartmentCostCentre nvarchar(20), @FirstName as varchar(50), @LastName as varchar(50), @InvoiceNumber varchar(50), @InvoiceDate varchar(100), @PODInfo as varchar(50) ")
        sbSQL.Append("DECLARE @CustomerRef1 nvarchar(30), @CustomerRef2 nvarchar(30), @Misc1 nvarchar(50), @Misc2 nvarchar(50), @SpecialInstructions nvarchar(50), @CneeTown varchar(50), @CneeCountryKey int, @CneeCountry varchar(50), @AgentKey int, @AgentName varchar(50), @WarehouseCutOffTime smalldatetime, @DespatchDate varchar(50), @ExternalSystemId nvarchar(50), @PODDate varchar(50) ")
        
        sbSQL.Append("SELECT [key], UserKey, StateId, Description INTO #temp2 FROM Consignment WHERE CustomerKey = 580 ")
        sbSQL.Append("DELETE FROM #temp2 WHERE StateId = 'CANCELLED' OR StateId IS NULL ")
        sbSQL.Append("DELETE FROM #temp2 WHERE Description = 'INTERNAL TRANSFER' ")
        sbSQL.Append("DELETE FROM #temp2 WHERE UserKey IN (5850, 11932, 11912, 11576) ")

        sbSQL.Append("DECLARE c CURSOR FOR ")
        sbSQL.Append("SELECT [key], UserKey FROM #temp2 ")
        If rbOrdersPlacedBy.Checked Then
            sbSQL.Append("WHERE UserKey = " & ddlOrdersOriginator.SelectedValue & " ")
        End If
        sbSQL.Append("OPEN c ")
        sbSQL.Append("FETCH NEXT FROM c INTO @ConsignmentKey, @BookedByKey ")
        sbSQL.Append("WHILE (@@FETCH_STATUS) = 0 ")
        sbSQL.Append("BEGIN ")
        sbSQL.Append("  SET @ConsignmentDate = (SELECT CreatedOn FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @ConsignmentWeight = (SELECT Weight FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @ConsignmentAmount = (SELECT CashOnDelAmount FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @ConsignmentNOP = (SELECT NOP FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @ConsignmentMode = (SELECT CneeTown FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  IF @ConsignmentMode LIKE '%mail%' SET @ConsignmentMode = 'Mail' ELSE SET @ConsignmentMode = 'Courier' ")
        sbSQL.Append("  SET @PODInfo = (SELECT ISNULL(PODDate,'') + ' ' + ISNULL(PODTime,'') + ' ' + ISNULL(PODName,'') FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @UserId = (SELECT UserId FROM UserProfile WHERE [key] = @BookedByKey) ")
        sbSQL.Append("  SET @FirstName = (SELECT FirstName FROM UserProfile WHERE [key] = @BookedByKey) ")
        sbSQL.Append("  SET @LastName = (SELECT LastName FROM UserProfile WHERE [key] = @BookedByKey) ")
        sbSQL.Append("  SET @UserDepartmentCostCentre = (SELECT dbo.NoOrg(Department) FROM UserProfile WHERE [key] = @BookedByKey) ")
        sbSQL.Append("  SET @InvoiceNumber = 0 ")
        sbSQL.Append("  SET @InvoiceDate = 0 ")
        sbSQL.Append("  SET @CustomerRef1 = (SELECT CustomerRef1 FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @CustomerRef2 = (SELECT CustomerRef2 FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @Misc1 = (SELECT Misc1 FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @Misc2 = (SELECT Misc2 FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @SpecialInstructions = (SELECT SpecialInstructions FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @CneeTown = (SELECT CneeTown FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @CneeCountryKey = (SELECT CneeCountryKey FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @CneeCountry = (SELECT CountryName FROM Country WHERE CountryKey = @CneeCountryKey) ")
        sbSQL.Append("  SET @AgentKey = 0 ")
        sbSQL.Append("  SET @AgentName = '' ")
        sbSQL.Append("  SET @WarehouseCutOffTime = (SELECT WarehouseCutOffTime FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  SET @DespatchDate = '' ")
        sbSQL.Append("  IF NOT @WarehouseCutOffTime IS NULL SET @DespatchDate = (SELECT CAST(@WarehouseCutOffTime AS varchar(50))) ")
        sbSQL.Append("  SET @ExternalSystemId = 0 ")
        sbSQL.Append("  SET @PODDate = (SELECT ISNULL(PODDate,'') FROM Consignment WHERE [key] = @ConsignmentKey) ")
        sbSQL.Append("  INSERT INTO #temp (ConsignmentDate, AWB, ConsignmentWeight, ConsignmentAmount, ConsignmentNOP, ConsignmentMode, CustomerRef1, CustomerRef2, Misc1, Misc2, SpecialInstructions, ProductCode, ProductDate, ProductDescription, ItemsOut, ProductCategory, UnitWeightGrams, UnitValue, UserId, UserDepartmentCostCentre, Consignee, CneeTown, CneeCountry, AgentName, DespatchDate, ExternalSystemId, InvoiceNumber, InvoiceDate, PODInfo, PODDate) ")
        sbSQL.Append("  SELECT @ConsignmentDate, CAST(@ConsignmentKey AS varchar(10)), @ConsignmentWeight, CashOnDelAmount, @ConsignmentNOP, ")
        sbSQL.Append("        @ConsignmentMode, @CustomerRef1, @CustomerRef2, @Misc1, @Misc2, @SpecialInstructions, lp.ProductCode, ")
        sbSQL.Append("        lp.ProductDate, lp.ProductDescription, lm.ItemsOut, ISNULL(lp.ProductCategory,''), ISNULL(lp.UnitWeightGrams,0), ")
        sbSQL.Append("        lp.UnitValue, ISNULL(@UserId,'') + ' (' + ISNULL(@FirstName,'') + ' ' + ISNULL(@LastName,'') + ')', ")
        sbSQL.Append("        ISNULL(@UserDepartmentCostCentre,'NO ORG'), ")
        sbSQL.Append("        ISNULL(CneeName,'') + ' ' + ISNULL(CneeAddr1,'') + ' ' + ISNULL(CneeAddr2,'') + ' ' + ISNULL(CneeAddr3,'') + ' ' + ISNULL(CneeTown,'') + ' ' + ISNULL(CneeState,'') + ' ' + ISNULL(CneePostcode,''), ")
        sbSQL.Append("        @CneeTown, @CneeCountry, @AgentName, @DespatchDate, @ExternalSystemId, @InvoiceNumber, @InvoiceDate, ")
        sbSQL.Append("        @PODInfo, @PODDate ")
        sbSQL.Append("  FROM LogisticMovement lm ")
        sbSQL.Append("  INNER JOIN LogisticProduct lp ")
        sbSQL.Append("  ON lp.LogisticProductKey = lm.LogisticProductKey ")
        sbSQL.Append("  INNER JOIN Consignment c ")
        sbSQL.Append("  ON lm.ConsignmentKey = c.[key] ")
        sbSQL.Append("  WHERE lm.ConsignmentKey = @ConsignmentKey ")
        sbSQL.Append("  FETCH NEXT FROM c INTO @ConsignmentKey, @BookedByKey ")
        sbSQL.Append("END ")
        sbSQL.Append("CLOSE c DEALLOCATE c ")
        sbSQL.Append("SELECT SUBSTRING(CONVERT(varchar(24), ConsignmentDate, 113), 1, 17) 'Cons Date', AWB, ConsignmentWeight 'Cons Wt', ConsignmentAmount 'Cons Cost', ConsignmentNOP 'NOP', ConsignmentMode 'Mail / Courier', CustomerRef1 + CustomerRef2 + Misc1 + Misc2 'Cust Ref', SpecialInstructions 'Spcl Instrs', ProductCode 'Prod Code', ProductDate 'Org', ProductDescription 'Descr', ItemsOut '# Items', UnitWeightGrams 'Unit Wt (gms)', UnitValue 'Cost Price', UserId 'User Id', UserDepartmentCostCentre 'User Dept/CC', Consignee, CneeTown 'Town', DespatchDate 'Dsptch Date', PODInfo 'POD Info', PODDate 'POD Date' FROM #temp ")
        gsSQL = sbSQL.ToString
        psDownloadQuery = gsSQL
        Dim oDataTable As DataTable
        If rbOrdersAll.Checked Then
            WebMsgBox.Show("Sorry, there is too much data to show all records.\nHowever you can download the full list by clicking the 'download to excel' link ")
        Else
            oDataTable = ExecuteQueryToDataTable(gsSQL)
            gvData.DataSource = oDataTable
            gvData.DataBind()
            pnlData.Visible = True
        End If
        pnlDownload.Visible = True
    End Sub

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title>NHS PIP SCR - Dashboard v1.3 (16APR10)</title>
    <style type="text/css">
        body
        {
            font-family: Verdana;
            font-size: xx-small;
            background-color: #E6FDF0;
        }
        .style2
        {
            color: #FF0000;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <table style="width: 100%">
        <tr>
            <td style="width: 80%">
                &nbsp;
                <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small">NHS PIP - Materials Ordering - Dashboard</asp:Label>
                <br />
                <br />
                <asp:Menu ID="Menu1" runat="server" OnMenuItemClick="Menu1_MenuItemClick" ItemWrap="True"
                    Orientation="Horizontal" Font-Size="Small" Width="95%">
                    <DynamicHoverStyle BackColor="#CCCCFF" />
                    <DynamicMenuStyle BackColor="#E6FDF0" />
                    <Items>
                        <asp:MenuItem Text="Stock&nbsp;allocation" Value="Stock allocation">
                            <asp:MenuItem Text="Master products" Value="Generic"/>
                            <asp:MenuItem Text="Assigned products by Organisation" Value="AssignedByOrganisation"/>
                            <asp:MenuItem Text="Assigned products by creation date" Value="AssignedByCreationDate"/>
                            <asp:MenuItem Text="Assigned products awaiting stock allocation, by Organisation" Value="AwaitingAllocation"/>
                        </asp:MenuItem>
                        <asp:MenuItem Text="Reservations,&nbsp;orders,&nbsp;back&nbsp;orders,&nbsp;goods&nbsp;in" Value="Stock reservations">
                            <asp:MenuItem Text="Reservations" Value="Reservations"/>
                            <asp:MenuItem Text="Orders" Value="Orders"/>
                            <asp:MenuItem Text="Back orders" Value="BackOrders"/>
                            <asp:MenuItem Text="Goods in" Value="GoodsIn"/>
                            <asp:MenuItem Text="Available&nbsp;for&nbsp;distribution" Value="AvailableForDistribution"/>
                        </asp:MenuItem>
                        <asp:MenuItem Text="Organisations" Value="Organisations">
                            <asp:MenuItem Text="Search Organisations" Value="SearchOrgs"/>
                            <asp:MenuItem Text="Organisations with no assigned products" Value="NoProducts"/>
                            <asp:MenuItem Text="Organisations with no assigned user accounts" Value="NoAccounts"/>
                            <asp:MenuItem Text="Organisation list" Value="OrganisationList"/>
                        </asp:MenuItem>
                        <asp:MenuItem Text="Maintenance" Value="Maintenance">
                            <asp:MenuItem Text="Activity log" Value="Activity"/>
                            <asp:MenuItem Text="Product visibility" Value="ProductVisibility"/>
                            <asp:MenuItem Text="User accounts" Value="Accounts"/>
                            <asp:MenuItem Text="Cross-linking consistency check" Value="ConsistencyCheck"/>
                        </asp:MenuItem>
                        <asp:MenuItem Text="About" Value="About">
                            <asp:MenuItem Text="Change history" Value="changehistory"/>
                            <asp:MenuItem Text="Help" Value="Help"/>
                        </asp:MenuItem>
                    </Items>
                </asp:Menu>
            </td>
            <td style="width: 20%" align="right">
                <img alt="" src="http://www.sprintexpress.co.uk/nhspip/images/logos/nhs_logo.jpg" />
            </td>
        </tr>
    </table>
    <asp:Panel ID="pnlIntro" runat="server" Width="100%">
        <b>&nbsp;Introduction<br />
        </b>
        <br />
        &nbsp;Choose the report(s) you require from the menu above.<br />
        <br />
        &nbsp;Note that the <b>Stock reservations</b> reports will not be available until the
        web form goes live.<br />
        <br />
        <b>&nbsp;Change history<br />
            <br />
        &nbsp;</b>18MAR10 - initial version released for evaluation<br />
        &nbsp;19MAR10 - added further menu items to Maintenance menu; added About menu item 
        (shows this page); added download to Excel button<br />
        &nbsp;25MAR10 - fixed problem with some reports not being available to download; 
        addressed menu display issues<br />
        &nbsp;27MAR10 - added reporting for Stock Reservations and Activity Log; expanded PCT 
        reporting to cover all organisations<br />
        &nbsp;13APR10 - rearranged some menu items; notes can now be added to back orders, eg 
        to record when a print order is placed, expected, qty ordered, etc.; Added Goods 
        In report<br />
        &nbsp;16APR10 - added orders report<br />
        <br />
        <b>Reports &amp; functionality still to be added<br />
        </b>
        <br />
        1. Report showing master products with backlog that have quantity greater than 
        minimum stock level<br />
        2. Tool to generate picks to allocate product from master product to ring-fenced 
        products awaiting top up, email users awaiting top up<br />
        3. Historic orders report<br />
        <br />
        &nbsp;Please email questions, problem reports and enhancement requests to
        <a href="mailto:m.quinn@sprintexpress.co.uk">m.quinn@sprintexpress.co.uk</a>, marking 
        your message &quot;<b><i>NHS PIP Dashboard</i></b>&quot;.
        Thank you.<br />
    </asp:Panel>
    <asp:Panel ID="pnlGeneric" runat="server" Width="100%">
        <b>&nbsp;Generic Products<br />
        </b>
        <br />
        &nbsp;Lists the &#39;master&#39; products, from which stock is allocated to individual 
        organisations.<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAssignedByOrganisation" runat="server" Width="100%">
        <b>&nbsp; Assigned Products by organisation<br />
        </b>
        <br />
        &nbsp; Lists, for each organisation the products assigned to that organisation. Some 
        products may not yet have had stock allocated to them.<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAssignedByCreationDate" runat="server" Width="100%">
        <b>&nbsp;Assigned Products by Creation Date<br />
        </b>
        <br />
        &nbsp;Lists, by date of product creation, all products assigned to organisations. 
        Some products may not yet have had stock allocated to them.<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAwaitingAllocation" runat="server" Width="100%">
        <b>&nbsp;Awaiting Allocation<br />
        </b>
        <br />
        &nbsp;Lists, for each organisation that has requested stock, the products assigned to 
        that organisation that do not currently have stock.<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlReservations" runat="server" Width="100%">
        <b>&nbsp;Stock Reservations<br />
        </b>
        <br />
        &nbsp;Lists the stock reservations that have been received<br />
        <br />
        &nbsp;<asp:CheckBox ID="cbStockReservationsLimitByDate" runat="server" 
            AutoPostBack="True" Text="Limit by date:" OnCheckedChanged="cbStockReservationsLimitByDate_CheckedChanged" />
        &nbsp; FROM:
        <asp:TextBox ID="tbStockReservationFrom" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="80px"/>
        &nbsp;<a ID="imgCalendarButton3" runat="server" href="javascript:;" 
            onclick="window.open('../PopupCalendar4.aspx?textbox=tbStockReservationFrom','cal','width=300,height=305,left=270,top=180')" 
            visible="true"><img ID="Img3" runat="server" alt="" border="0" 
            ie:visible="true" src="images/SmallCalendar.gif" visible="false" /></a> (eg 
        1-jun-10) TO:
        <asp:TextBox ID="tbStockReservationTo" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="80px"/>
        &nbsp;<a ID="imgCalendarButton4" runat="server" href="javascript:;" 
            onclick="window.open('../PopupCalendar4.aspx?textbox=tbStockReservationTo','cal','width=300,height=305,left=270,top=180')" 
            visible="true"><img ID="Img4" runat="server" alt="" border="0" 
            ie:visible="true" src="images/SmallCalendar.gif" visible="false" /></a> (eg 
        31-jun-10)&nbsp;
        <asp:LinkButton ID="lnkbtnStockReservationDateLast7Days" runat="server" OnClick="lnkbtnStockReservationDateLast7Days_Click">last 7 days</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnStockReservationDateLast30Days" runat="server" OnClick="lnkbtnStockReservationDateLast30Days_Click">last 30 days</asp:LinkButton>
        <br />
        <br />
        &nbsp;<asp:RadioButton ID="rbStockReservationsFilterByOrganisation" 
            runat="server" AutoPostBack="true" GroupName="StockReservationsFilter" 
            OnCheckedChanged="rbStockReservationsFilterByOrganisation_CheckedChanged" 
            Text="Filter by organisation:" />
        <asp:DropDownList ID="ddlStockReservationsOrganisation" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        &nbsp;<asp:RadioButton ID="rbStockReservationsNoFiltering" runat="server" 
            AutoPostBack="true" Checked="True" GroupName="StockReservationsFilter" 
            OnCheckedChanged="rbStockReservationsNoFiltering_CheckedChanged" 
            Text="No filtering" />
        <br />
        &nbsp; Show: reservations where&nbsp;<asp:CheckBox ID="cbStockReservationsFullyAssigned" runat="server" 
            Checked="True" Text="stock fully assigned" />
        &nbsp;<asp:CheckBox ID="cbStockReservationsPartiallyAssigned" runat="server" 
            Checked="True" Text="stock partially assigned" />
        &nbsp;<asp:CheckBox ID="cbStockReservationsUnassigned" runat="server" Checked="True" 
            Text="stock unassigned" />
        <br />
        <br />
        &nbsp;<asp:Button ID="btnStockReservationsGo" runat="server" Text="go" 
            Width="200px" OnClick="btnStockReservationsGo_Click" />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlOrders" runat="server" Width="100%">
        <b>&nbsp;Orders</b><br />
        <br />
        &nbsp;Shows orders placed using the NHS PIP draw down account,&nbsp;<b>one product per 
        line</b> (ie a single order containing multiple products will occupy as many 
        lines as there are products in that order).<br />
        <br />
        &nbsp;You can list all orders, or orders for a specific organisation/user.<br />
        <br />
        &nbsp;NOTE: Sprint-generated orders (eg internal transfers) are not included in this 
        list.<br />
        <br />
        <asp:RadioButton ID="rbOrdersAll" runat="server" AutoPostBack="True" 
            Checked="True" Font-Names="Verdana" Font-Size="XX-Small" GroupName="orders" 
            OnCheckedChanged="rbOrdersAll_CheckedChanged" Text="show all orders" />
        <asp:RadioButton ID="rbOrdersPlacedBy" runat="server" AutoPostBack="True" 
            Font-Names="Verdana" Font-Size="XX-Small" GroupName="orders" 
            oncheckedchanged="rbOrdersPlacedBy_CheckedChanged" 
            Text="show orders placed by" />
        &nbsp;<asp:DropDownList ID="ddlOrdersOriginator" runat="server" Enabled="False" 
            Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        &nbsp;<asp:Button ID="btnOrderGo" runat="server" Text="go" Width="200px" onclick="btnOrderGo_Click" />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlGoodsIn" runat="server" Width="100%">
        <b>&nbsp;Goods In<br />
        <br />
        &nbsp;</b>Goods in activity to show:&nbsp;
        <asp:RadioButton ID="rbGoodsInMostRecent" runat="server" Checked="True" 
            GroupName="GoodsInNumberOfEvents" Text="most recent event" />
        &nbsp;
        <asp:RadioButton ID="rbGoodsInAll" runat="server" 
            GroupName="GoodsInNumberOfEvents" Text="all events" />
        <br />
        <br />
        &nbsp;Products to show:&nbsp;<asp:RadioButton ID="rbGoodsInProductsWithBackOrders" 
            runat="server" Checked="True" GroupName="GoodsInProducts" 
            Text="products with back orders only" />
        &nbsp;
        <asp:RadioButton ID="rbGoodsInProductsAll" runat="server" 
            GroupName="GoodsInProducts" Text="all products" />
        <br />
        <br />
        &nbsp;<asp:Button ID="btnGoodsIn" runat="server" Text="go" Width="194px" OnClick="btnGoodsIn_Click" />
        <br />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAvailableForDistribution" runat="server" Width="100%">
        <b>&nbsp;&nbsp;Available for distribution</b>
        <br />
        <br />
        &nbsp;Shows MASTER products with back orders where the available quantity is more 
        than minimum stock level, and therefore available for topping up one or more PER 
        ORGANISATION products.<br />
        <br />
        Click the <b>distribute!</b> button to display a screen that allows you to 
        generate picks to transfer stock from the MASTER product to selected PER 
        ORGANISATION products. You can specify the amount to transfer to each PER 
        ORGANISATION product.<br />
        <br />
        <br />
        <asp:GridView ID="gvAvailableForDistribution" runat="server" CellPadding="2" 
            Width="100%" AutoGenerateColumns="False">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="btnDistribute" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' Text="distribute!" OnClick="btnDistribute_Click" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Product" HeaderText="Product" ReadOnly="True" 
                    SortExpression="Product" />
                <asp:BoundField DataField="Description" HeaderText="Description" 
                    ReadOnly="True" SortExpression="Description" />
                <asp:BoundField DataField="MinimumStockLevel" HeaderText="Min stock level" 
                    ReadOnly="True" SortExpression="MinimumStockLevel" />
                <asp:BoundField DataField="Qty" HeaderText="Qty available" ReadOnly="True" 
                    SortExpression="Qty" />
            </Columns>
            <EmptyDataTemplate>
                <i>no products available for distribution</i>
            </EmptyDataTemplate>
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlDistribute" runat="server" Width="100%">
        &nbsp;<b>Distribute products</b>
        <br />
        <br />
        &nbsp;Transfers stock from MASTER product to selected PER ORGANISATION products with 
        backlogs. The system suggests the amount to transfer. You can adjust this 
        amount. If the transfer amount is greater than or equal to the back order 
        quantity requested, the back order is closed, but you can override this using 
        the <b>close back order</b> check box. Click the <b>recalculate</b> button to 
        check the quantities to transfer and show the expected MASTER product stock 
        quantity remaining after the distribution.&nbsp; Click the start <b>distribution 
        button</b> to initiate the distribution. The system may take a few seconds to 
        process the request, generate the picks and adjust the stock levels.<br />
        <br />
        <b>&nbsp;Product:</b>
        <asp:Label ID="lblDistributeProduct" runat="server"></asp:Label>
        <br />
        <b>&nbsp;Available stock:</b>
        <asp:Label ID="lblDistributeAvailableStockQty" runat="server"></asp:Label>
        <br />
        <br />
        <asp:GridView ID="gvDistribute" runat="server" CellPadding="2" Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvDistribute_RowDataBound">
            <Columns>
                <asp:BoundField DataField="ProductDate" HeaderText="Organisation" ReadOnly="True" />
                <asp:BoundField DataField="Quantity" HeaderText="Current stock qty" ReadOnly="True" />
                <asp:TemplateField HeaderText="Back order qty requested">
                    <ItemTemplate>
                        <asp:Label ID="lblBackLogQty" runat="server" Text='<%# Bind("BackLogQty") %>'/>
                        <asp:LinkButton ID="lnkbtnTransfer" CommandArgument='<%# Container.DataItemIndex %>' runat="server" OnClick="lnkbtnTransfer_Click">transfer</asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Qty to transfer">
                    <ItemTemplate>
                        &nbsp;
                        <asp:TextBox ID="tbQtyToTransfer" runat="server" Width="50" Font-Names="Verdana" Font-Size="XX-Small" >0</asp:TextBox>
                        <asp:RangeValidator ID="rvQtyToTransfer" runat="server" 
                            ControlToValidate="tbQtyToTransfer" ErrorMessage="reqd" MaximumValue="1000000" 
                            MinimumValue="0" Type="Integer"/><asp:RequiredFieldValidator ID="rfvQtyToTransfer" runat="server" ControlToValidate="tbQtyToTransfer" ErrorMessage="reqd"/>
                        <asp:CheckBox ID="cbCloseBackOrder" runat="server" Text="close back order" Font-Names="Verdana" Font-Size="XX-Small" />
                        <asp:HiddenField ID="hidLogisticProductKey" runat="server" Value='<%# Container.DataItem("RingFencedProductKey")%>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <EmptyDataTemplate>
                <i>no products available for distribution</i>
            </EmptyDataTemplate>
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
        <br />
        &nbsp;Stock level after distribution:
        <asp:Label ID="lblDistributeStockLevelAfterDistribution" runat="server"></asp:Label>
        &nbsp;<asp:LinkButton ID="lnkbtnRecalculate" runat="server" OnClick="lnkbtnRecalculate_Click">recalculate</asp:LinkButton>
        &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="btnStartDistribution" runat="server" Text="start distribution" Width="200px" OnClick="btnStartDistribution_Click" />
    </asp:Panel>
    <asp:Panel ID="pnlBackOrders" runat="server" Width="100%">
        <b>&nbsp; Back orders<br />
        <br />
        &nbsp;
        </b>
        Shows in the MASTER PRODUCT listing the total back order quantity (if any) for 
        each MASTER product, in the PER ORGANISATION listing the back order quantity (if 
        any) for each product created for&nbsp; an organisation to draw down.&nbsp; The total for 
        any given MASTER product in the top listing is the sum of all the back order 
        quantities for PER ORGANISATION products of that product type in the bottom 
        listing.<br />
        <br />
        &nbsp;Click a button in the MASTER PRODUCT&nbsp; listing to remove all back order 
        quantities for that MASTER product. Click a button in the PER ORGANISATION&nbsp; 
        listing to remove the back order quantity from that PER ORGANISATION product.<br />
        <br />
        &nbsp;Click <b>add note</b>&nbsp; to add a note to a MASTER product or a PER ORGANISATION 
        product. Typically you use notes to record details of print orders, expected 
        deliveries, additional order details, etc.<br />
        <br />
        <br />
        
        &nbsp;
        
    </asp:Panel>
    <asp:Panel ID="pnlSearchOrgs" runat="server" Width="100%">
        <b>&nbsp;Search organisations<br />
        </b>
        <br />
        &nbsp;Search by organisation code or partial organisation name:
        <asp:TextBox ID="tbSearchOrgs" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="150px"></asp:TextBox>
        &nbsp;<asp:Button ID="btnSearchGo" runat="server" Text="go" OnClick="btnSearchGo_Click" />
        <br />
        <br />
        &nbsp;
    </asp:Panel>
    <asp:Panel ID="pnlNoProducts" runat="server" Width="100%">
        <b>&nbsp;Organisations with no assigned products<br />
        </b>
        <br />
        &nbsp;Lists the organisations that have no products assigned to them<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlNoAccounts" runat="server" Width="100%">
        <b>Organisations with no user accounts<br />
        </b>
        <br />
        Lists the organisations for which no user accounts exist<br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlActivity" runat="server" Width="100%">
        <b>&nbsp;Activity Log<br />
        </b>
        <br />
        &nbsp;Displays the detailed activity log for stock reservations made from the web 
        form<br />
        <br />
        &nbsp;<asp:CheckBox ID="cbActivityLogLimitByDate" runat="server" 
            Text="Limit by date:" AutoPostBack="True" OnCheckedChanged="cbActivityLogLimitByDate_CheckedChanged" />
        &nbsp; FROM:
        <asp:TextBox ID="tbActivityLogFrom" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="80px"></asp:TextBox>
        &nbsp;<a ID="imgCalendarButton1" runat="server" href="javascript:;" 
            onclick="window.open('../PopupCalendar4.aspx?textbox=tbActivityLogFrom','cal','width=300,height=305,left=270,top=180')" 
            visible="true"><img ID="Img1" runat="server" alt="" border="0" 
            ie:visible="true" src="images/SmallCalendar.gif" visible="false" /></a> (eg 
        1-jun-10) TO:
        <asp:TextBox ID="tbActivityLogTo" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="80px"></asp:TextBox>
        &nbsp;<a ID="imgCalendarButton2" runat="server" href="javascript:;" 
            onclick="window.open('../PopupCalendar4.aspx?textbox=tbActivityLogTo','cal','width=300,height=305,left=270,top=180')" 
            visible="true"><img ID="Img2" runat="server" alt="" border="0" 
            ie:visible="true" src="images/SmallCalendar.gif" visible="false" /></a> (eg 
        31-jun-10)&nbsp;
        <asp:LinkButton ID="lnkbtnActivityLogDateLast7Days" runat="server" onclick="lnkbtnActivityLogDateLast7Days_Click">last 7 days</asp:LinkButton>
        &nbsp;<asp:LinkButton ID="lnkbtnActivityLogDateLast30Days" runat="server" onclick="lnkbtnActivityLogDateLast30Days_Click">last 30 days</asp:LinkButton>
        <br />
        <br />
        <asp:RadioButton ID="rbActivityLogFilterByStockAdditionFailures" runat="server" AutoPostBack="true" GroupName="ActivityLogFilter" Text="Filter - SHOW STOCK ADDITION FAILURES" OnCheckedChanged="rbActivityLogFilterByStockAdditionFailures_CheckedChanged" />
        &nbsp;(shows partial and complete failures to auto-assign stock from MASTER to 
        ring-fenced products)<br />
        <asp:RadioButton ID="rbActivityLogFilterByOrderNo" runat="server" AutoPostBack="true" GroupName="ActivityLogFilter" Text="Filter by Order #:" OnCheckedChanged="rbActivityLogFilterByOrderNo_CheckedChanged" />
        &nbsp;<asp:TextBox ID="tbActivityLogOrderNo" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" Width="120px" />
        <br />
        <asp:RadioButton ID="rbActivityLogFilterByProduct" runat="server" 
            AutoPostBack="true" GroupName="ActivityLogFilter" 
            OnCheckedChanged="rbActivityLogFilterByProduct_CheckedChanged" 
            Text="Filter by Product:" />
        &nbsp;<asp:DropDownList ID="ddlActivityLogFilterProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small" />
        <br />
        <asp:RadioButton ID="rbActivityLogActivityLogText" runat="server" 
            GroupName="ActivityLogFilter" Text="Activity log text" 
            OnCheckedChanged="rbActivityLogActivityLogText_CheckedChanged" 
            AutoPostBack="True" />
        :
        <asp:TextBox ID="tbActivityLogText" runat="server" Font-Names="Verdana" 
            Font-Size="XX-Small" Width="227px" />
        <br />
        <asp:RadioButton ID="rbActivityLogNoFiltering" runat="server" AutoPostBack="true" Checked="True" GroupName="ActivityLogFilter" Text="No filtering" OnCheckedChanged="rbActivityLogNoFiltering_CheckedChanged" />
        <br />
        <br />
        &nbsp;<asp:Button ID="btnActivityLogGo" runat="server" Text="go" Width="194px" onclick="btnActivityLogGo_Click" />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlOrganisationList" runat="server" Width="100%">
        <b>&nbsp;Organisation List<br />
        </b>
        <br />
        &nbsp;Displays the list of organisations, organisation abbreviations and organisation 
        codes used within the system. Abbreviations are used on products assigned to 
        organisations to show the assignment, and are used internally to link user 
        accounts with sets of assigned products. Virtual organisations are organisations 
        such as &quot;Hambleton &amp; Richmondshire&quot; used internally to map users to product sets 
        where the actual organisation intended is unclear, or organisations such as 
        &#39;Sussex&quot; where a single user account will be ordering for multiple real 
        organisations.<br />
        <br />
        &nbsp;<asp:RadioButton ID="rbShowRealOrganisations" runat="server" 
            Checked="True" AutoPostBack="true"
            OnCheckedChanged="rbShowRealOrgs_CheckedChanged" GroupName="Orgs" 
            Text="show real organisations" />
        &nbsp;<asp:RadioButton ID="rbShowVirtualOrganisatios" runat="server" 
            GroupName="Orgs" AutoPostBack="true"
            OnCheckedChanged="rbShowVirtualOrgs_CheckedChanged" 
            Text="show virtual organisations" />
        &nbsp;
        <asp:DropDownList ID="ddlOrganisationFilter" runat="server" AutoPostBack="True" 
            OnSelectedIndexChanged="ddlOrganisationFilter_SelectedIndexChanged" 
            Font-Names="Verdana" Font-Size="XX-Small">
            <asp:ListItem Value="1 = 1">show all organisations</asp:ListItem>
            <asp:ListItem Value="Type = 'PCT'">show PCTs</asp:ListItem>
            <asp:ListItem Value="Type = 'TRUST'">show trusts</asp:ListItem>
            <asp:ListItem Value="Type = 'CARE'">show care trusts</asp:ListItem>
            <asp:ListItem Value="Type = 'RO'">show regional offices</asp:ListItem>
            <asp:ListItem Value="Type = 'SHA'">show SHAs</asp:ListItem>
        </asp:DropDownList>
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlProductVisibility" runat="server" Width="100%">
        <b>&nbsp;Product Visibility<br />
        </b>
        <br />
        &nbsp;Lists the products visible to each user account.<br />
        &nbsp;<asp:Label ID="lblLegendMismatches" runat="server" Font-Bold="True" 
            ForeColor="Red" Text="MISMATCHES" Visible="False"></asp:Label>
        <asp:Label ID="lblVisibilityMismatches" runat="server" Font-Bold="True" 
            ForeColor="Red" 
            Text="There are possible mismatches  Click the button to view these mismatches." 
            Visible="False"></asp:Label>
        &nbsp;<asp:Button ID="btnVisibilityMismatches" runat="server" Text="view visibility mismatches"
            OnClick="btnVisibilityMismatches_Click" Visible="False" />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlAccounts" runat="server" Width="100%">
        <b>&nbsp;Accounts<br />
        </b>
        <br />
        &nbsp;Lists the accounts created for each organisation.<br />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlConsistencyCheck" runat="server" Width="100%">
        <b>Consistency Check<br />
        </b>
        <br />
        Runs a consistency check across all products and user accounts, to ensure the 
        cross linked organisation data is correct.<br />
        <br />
        <span class="style2"><b>If any products or user accounts are displayed, this indicates
            a discrepancy in the cross linking. Please inform Sprint IT immediately.</b></span><br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlDownload" runat="server" Width="100%">
    <asp:LinkButton ID="lnkbtnDownload" runat="server" OnClick="lnkbtnDownload_Click">download to excel</asp:LinkButton>
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%">
        <b>Help<br />
        </b>
        <br />
        <i><b>1.&nbsp; How the NHS PIP web form works<br />
        </b></i>
        <br />
        The web form lets designated NHS staff notify CFH of their requirements for CRS 
        materials for use in their organisation.&nbsp; Some users may be ordering for more 
        than one organisation, in which
        <br />
        case their &#39;owning&#39; organisation is used as a reference. Typically the web form 
        user will register for a draw down account on the Sprint online stock system and 
        select quantities of the materials they require.<br />
        <br />
        When the order is submitted, the system creates the requested user account and 
        notifiies the user by email. Then for each product requested, the system creates 
        a new &#39;ring fenced&#39; product for that locality (eg PCT, Trust).&nbsp; The newly 
        created product has the same product code as the &#39;master&#39; product but is 
        differentiated by the Product Date, which shows the abbreviation or code for the 
        user&#39;s organisation.<br />
        <br />
        The system then tries to transfer the requested amount of stock from the master 
        product to the newly created product. It will always leave at least the &#39;Minimum 
        Stock Level&#39; amount in the master product. Stock above the minimum stock level 
        is referred to heareafter as &#39;surplus stock&#39;.&nbsp; This transfer may succeed (there 
        is enough surplus stock to transfer the entire quantity requested), partially 
        succeed (some but not all of the quantity or surplus stock requested is 
        transferred) or fail (no surplus stock is available for transfer).<br />
        <br />
        <br />
    </asp:Panel>
    <asp:Panel ID="pnlData" runat="server" Width="100%">
        <asp:GridView ID="gvData" runat="server" CellPadding="2" Width="100%">
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlData2" runat="server" Width="100%">
        <br />
        <br />
        <asp:GridView ID="gvData2" runat="server" CellPadding="2" Width="100%">
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
    </asp:Panel>
    <asp:Panel ID="pnlBackOrdersData" runat="server" Width="100%">
        <b>Back orders - BY MASTER PRODUCT (shows sum of per-organisation product back 
        orders for each master product)<br />
        <br />
        </b>&nbsp;<asp:GridView ID="gvBackOrdersSummary" runat="server" CellPadding="2" 
            Width="100%" AutoGenerateColumns="False">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="btnClearAllOrganisationProductBackOrders" 
                            CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' runat="server" 
                            Text="clear all back orders" OnClientClick='return confirm("Are you sure you want to clear all back orders for this product?");'
                            onclick="btnClearAllOrganisationProductBackOrders_Click" />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ProductCode" HeaderText="Product" ReadOnly="True" SortExpression="Product" />
                <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="Description" />
                <asp:BoundField DataField="Qty" HeaderText="Back Order Qty" ReadOnly="True" SortExpression="Qty" />
                <asp:TemplateField HeaderText="Note">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnAddOrganisationProductBackOrderNote" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' OnClick="lnkbtnAddBackOrderNote_Click">add note</asp:LinkButton>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("Note") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
        <br />
        <b>Back orders - BY PER-ORGANISATION PRODUCT<br />
        <br />
        </b>
        <asp:GridView ID="gvBackOrdersDetail" runat="server" CellPadding="2" 
            Width="100%" AutoGenerateColumns="False">
            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:Button ID="btnClearOrganisationProductBackOrder" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' runat="server" 
                            Text="clear back order" OnClick="btnClearOrganisationProductBackOrder_Click" OnClientClick='return confirm("Are you sure you want to clear the back order quantity for this product?");' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Organisation" HeaderText="Organisation" ReadOnly="True" SortExpression="Organisation" />
                <asp:BoundField DataField="Product" HeaderText="Product" ReadOnly="True" SortExpression="Product" />
                <asp:BoundField DataField="Description" HeaderText="Description" ReadOnly="True" SortExpression="Description" />
                <asp:BoundField DataField="Qty" HeaderText="Back Order Qty" ReadOnly="True" SortExpression="Qty" />
                <asp:TemplateField HeaderText="Note">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnAddOrganisationProductBackOrderNote" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' OnClick="lnkbtnAddBackOrderNote_Click">add note</asp:LinkButton>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("Note") %>'/>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <AlternatingRowStyle BackColor="#FFFFCC" />
        </asp:GridView>
        <br />
        <b>&nbsp;Add a back order<br />
        <br />
        &nbsp;</b>You can add (or edit) a back order manually using the facility below. You 
        can only add a back order for a product that already exists.<br />
        <br />
        &nbsp;Product:
        <asp:DropDownList ID="ddlAddBackOrderProduct" runat="server" 
            Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlAddBackOrderProduct_SelectedIndexChanged"/>
        &nbsp;Organisation:
        <asp:DropDownList ID="ddlAddBackOrderOrganisations" runat="server" 
            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" 
            OnSelectedIndexChanged="ddlAddBackOrderOrganisations_SelectedIndexChanged" />
        <br />
        Qty:<asp:TextBox ID="tbAddBackOrderQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="80px"/>
        &nbsp;<asp:RangeValidator ID="AddBackOrderQty" runat="server" 
            ControlToValidate="tbAddBackOrderQty" ErrorMessage="number reqd" 
            MaximumValue="1000000" MinimumValue="0" Type="Integer"></asp:RangeValidator>
        &nbsp;<asp:Button ID="btnAddBackOrderGo" runat="server" 
            OnClick="btnAddBackOrderGo_Click" Text="go" Width="131px" />
    </asp:Panel>
    <asp:Panel ID="pnlAddNote" runat="server" Width="100%">
        <br />
        <b>Add Note to
        <asp:Label ID="lblProductNote" runat="server"></asp:Label>
        <br />
        <br />
        </b>
        <br />
        Note:
        <asp:TextBox ID="tbNote" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="600px"/>
        &nbsp;<asp:Button ID="btnSaveNote" runat="server" Text="save note" 
            OnClick="btnSaveNote_Click" Width="100px" />
        &nbsp;<asp:Button ID="btnCancelNote" runat="server" Text="cancel" OnClick="btnCancelNote_Click" />
        <br />
        <br />
        <b>PREVIOUS NOTES</b><br />
        <br />
        <asp:Label ID="lblProductCurrentNote" runat="server"/>
        <br />
    </asp:Panel>
    </form>
</body>
</html>
