<%@ Page Language="VB" Theme="AIMSDefault" MaintainScrollPositionOnPostback="true" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Drawing.Color" %>
<%@ import Namespace="System.Globalization" %>
<%@ import Namespace="System.Threading" %>

<script runat="server">

    Const COUNTRY_KEY_UK As Integer = 222
    Const CUSTOMER_LOVELLS As Integer = 663
    Const CUSTOMER_HARDIE As Int32 = 837
    Const CUSTOMER_HARDFR As Int32 = 849
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call PopulateUserProfileFields()
            Call SetUserPrivilegeVisibility()
            
            If ConfigLib.GetConfigItem_EnableAuthorisation Then
                Call GetPendingOrderAuthorisations()
                If gvAuthoriseOrder.Rows.Count > 0 Then
                    Call ShowAuthorisationsPanel()
                End If
            End If
            pnlCalendarManagement.Visible = ConfigLib.GetConfigItem_EnableCalendarManagement
            
            If IsHardie() Or IsHardFR() Then
                trHardieMonthlyCreditRemaining.Visible = True
                Call SetHardieCreditRemaining()
                If Session("UserType").ToString.ToLower.Contains("super") Then
                    trHardieManageCredit01.Visible = True
                End If
            End If
        End If
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)
        Call SetTitle()
        Call SetFeatureVisibility()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "My Profile"
    End Sub
    
    Protected Sub SetHardieCreditRemaining()
        Dim nHardieCreditRemaining As Double = GetJamesHardieMonthlyOrderRemainingAllowance()
        If nHardieCreditRemaining < 0 Then
            lblHardieCreditRemaining.Text = "no transactions yet, or no limit"
        Else
            lblHardieCreditRemaining.Text = nHardieCreditRemaining.ToString("C", CultureInfo.CurrentCulture)
        End If
    End Sub
    
    Protected Function GetJamesHardieMonthlyOrderRemainingAllowance() As Double   ' -1 = no checking, unlimited or no value defined
        GetJamesHardieMonthlyOrderRemainingAllowance = -1
        Dim dblMonthOrderValueToDate As Double = 0
        Dim dblUserMonthOrderLimit As Double = -1
        Dim dblDefaultMonthOrderLimit As Double = -1
        Dim bOverrideCheck As Boolean = False
        Call ExecuteQueryToDataTable("UPDATE ClientData_JamesHardie_MonthOrderValue SET MonthOrderValue = 0 WHERE UserKey = " & Session("UserKey") & " AND DATEPART(m, LastUpdatedOn) <> DATEPART(m, GETDATE())")
        Dim dtJamesHardie_MonthOrderValue As DataTable = ExecuteQueryToDataTable("SELECT MonthOrderValue, OverrideCheck FROM ClientData_JamesHardie_MonthOrderValue WHERE UserKey = " & Session("UserKey"))
        If dtJamesHardie_MonthOrderValue.Rows.Count = 1 Then
            dblMonthOrderValueToDate = dtJamesHardie_MonthOrderValue.Rows(0).Item("MonthOrderValue")
            bOverrideCheck = dtJamesHardie_MonthOrderValue.Rows(0).Item("OverrideCheck")
        End If
        If Not bOverrideCheck Then
            Dim dtJamesHardie_MonthOrderLimit As DataTable = ExecuteQueryToDataTable("SELECT MonthOrderLimit, UserKey FROM ClientData_JamesHardie_MonthOrderLimit WHERE UserKey IN (" & Session("UserKey") & ", 0)")
            For Each drJamesHardie_MonthOrderLimit As DataRow In dtJamesHardie_MonthOrderLimit.Rows
                If drJamesHardie_MonthOrderLimit("UserKey") = Session("UserKey") Then
                    dblUserMonthOrderLimit = drJamesHardie_MonthOrderLimit("MonthOrderLimit")
                ElseIf drJamesHardie_MonthOrderLimit("UserKey") = 0 Then
                    dblDefaultMonthOrderLimit = drJamesHardie_MonthOrderLimit("MonthOrderLimit")
                End If
            Next
            If dblUserMonthOrderLimit >= 0 Then
                GetJamesHardieMonthlyOrderRemainingAllowance = dblUserMonthOrderLimit - dblMonthOrderValueToDate
                If GetJamesHardieMonthlyOrderRemainingAllowance < 0 Then
                    GetJamesHardieMonthlyOrderRemainingAllowance = 0
                End If
            ElseIf dblDefaultMonthOrderLimit >= 0 Then
                GetJamesHardieMonthlyOrderRemainingAllowance = dblDefaultMonthOrderLimit - dblMonthOrderValueToDate
                If GetJamesHardieMonthlyOrderRemainingAllowance < 0 Then
                    GetJamesHardieMonthlyOrderRemainingAllowance = 0
                End If
            End If
        End If
    End Function

    Protected Function IsHardie() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHardie = IIf(gbSiteTypeDefined, gsSiteType = "hardie", nCustomerKey = CUSTOMER_HARDIE)
    End Function

    Protected Function IsHardFR() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHardFR = IIf(gbSiteTypeDefined, gsSiteType = "hardfr", nCustomerKey = CUSTOMER_HARDFR)
    End Function

    Protected Sub SetFeatureVisibility()
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
        pnlCalendarManagement.Visible = dr("CalendarManagement")
        tbCostCentre.Enabled = CBool(dr("UserCanChangeCostCentre"))
        
        trAuthorisationExempt.Visible = False
        Dim bOrderAuthorisation As Boolean = dr("OrderAuthorisation")
        Dim bProductAuthorisation As Boolean = dr("ProductAuthorisation")
        If bOrderAuthorisation Or bProductAuthorisation Then
            If Not UserMustAuthorise() Then
                trAuthorisationExempt.Visible = True
            End If
        End If
    End Sub

    Protected Sub PopulatePublicationStatus()
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_UserProfile_GetPublicationProfile", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParam.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("ReceivePublicationOrderAlerts")) Then
                cbReceivePublicationOrderAlerts.Checked = oDataReader("ReceivePublicationOrderAlerts")
            End If
            If Not IsDBNull(oDataReader("ReceiveProductInactivityAlerts")) Then
                cbReceiveProductInactivityAlerts.Checked = oDataReader("ReceiveProductInactivityAlerts")
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in PopulatePublicationStatus: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function IsLovells() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsLovells = IIf(gbSiteTypeDefined, gsSiteType = "lovells", nCustomerKey = CUSTOMER_LOVELLS)
    End Function
    
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
    
    Protected Sub HideAllPanels()
        pnlOrderDetail.Visible = False
        pnlAuthorisations.Visible = False
        pnlEvent.Visible = False
    End Sub
    
    Protected Sub ShowAuthorisationsPanel()
        HideAllPanels()
        pnlAuthorisations.Visible = True
    End Sub
    
    Protected Sub ShowOrderDetailPanel()
        HideAllPanels()
        pnlOrderDetail.Visible = True
    End Sub
    
    Protected Sub ShowEventPanel()
        HideAllPanels()
        pnlEvent.Visible = True
    End Sub
    
    Protected Sub SetUserPrivilegeVisibility()
        tblAlerts.Visible = False
        trReceiveOrderConfirmationAlerts.Visible = False
        trReceiveGoodsInAlerts.Visible = False
        trReceiveLowStockAlerts.Visible = False
        
        cbReceiveLowStockAlertsSuperUser.Visible = False
        cbReceiveLowStockAlertsProductOwner.Visible = False
        cbReceiveGoodsInAlertsSuperUser.Visible = False
        cbReceiveGoodsInAlertsProductOwner.Visible = False
        cbReceiveOrderConfirmationAlertsSuperUser.Visible = False
        cbReceiveOrderConfirmationAlertsProductOwner.Visible = False
        cbReceiveOrderConfirmationAlertsUser.Visible = False

        If Session("UserType").ToString.ToLower = "superuser" Then
            tblAlerts.Visible = True
            trReceiveOrderConfirmationAlerts.Visible = True
            trReceiveGoodsInAlerts.Visible = True
            trReceiveLowStockAlerts.Visible = True

            cbReceiveOrderConfirmationAlertsUser.Visible = True
            cbReceiveLowStockAlertsSuperUser.Visible = True
            cbReceiveGoodsInAlertsSuperUser.Visible = True
            cbReceiveOrderConfirmationAlertsSuperUser.Visible = True
        End If

        If Session("UserType").ToString.ToLower.Contains("owner") Then
            tblAlerts.Visible = True
            trReceiveOrderConfirmationAlerts.Visible = True
            trReceiveGoodsInAlerts.Visible = True
            trReceiveLowStockAlerts.Visible = True

            cbReceiveOrderConfirmationAlertsUser.Visible = True
            cbReceiveLowStockAlertsProductOwner.Visible = True
            cbReceiveGoodsInAlertsProductOwner.Visible = True
            
            tblProductCodeReservations.Visible = True
            ' Call GetProductCodeReservations()
        End If

        If Session("UserType").ToString.ToLower = "user" Then
            tblAlerts.Visible = True
            cbReceiveOrderConfirmationAlertsUser.Visible = True
        End If
        
        If IsLovells() And Session("UserType").ToString.ToLower = "user" Then
            trLovellsUser.Visible = True
            Call PopulatePublicationStatus()
        End If
    End Sub
    
    Protected Sub PopulateUserProfileFields()
        Dim nDefaultDestinationGABKey As Integer
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

            If Not IsDBNull(oDataReader("Department")) Then
                tbCostCentre.Text = oDataReader("Department")
            End If
            If Not IsDBNull(oDataReader("EmailAddr")) Then
                tbEmail.Text = oDataReader("EmailAddr")
            End If
            If Not IsDBNull(oDataReader("Telephone")) Then
                tbTelephone.Text = oDataReader("Telephone")
            End If


            If Not IsDBNull(oDataReader("StockBookingAlert")) Then
                If oDataReader("StockBookingAlert") Then
                    cbReceiveOrderConfirmationAlertsUser.Checked = True
                Else
                    cbReceiveOrderConfirmationAlertsUser.Checked = False
                End If
            Else
                cbReceiveOrderConfirmationAlertsUser.Checked = False
            End If

            If Not IsDBNull(oDataReader("StockBookingAlertAll")) Then
                If oDataReader("StockBookingAlertAll") Then
                    cbReceiveOrderConfirmationAlertsProductOwner.Checked = True
                    cbReceiveOrderConfirmationAlertsSuperUser.Checked = True
                Else
                    cbReceiveOrderConfirmationAlertsProductOwner.Checked = False
                    cbReceiveOrderConfirmationAlertsSuperUser.Checked = False
                End If
            Else
                cbReceiveOrderConfirmationAlertsProductOwner.Checked = False
                cbReceiveOrderConfirmationAlertsSuperUser.Checked = False
            End If

            If Not IsDBNull(oDataReader("StockArrivalAlert")) Then
                If oDataReader("StockArrivalAlert") Then
                    cbReceiveGoodsInAlertsProductOwner.Checked = True
                    cbReceiveGoodsInAlertsSuperUser.Checked = True
                Else
                    cbReceiveGoodsInAlertsProductOwner.Checked = False
                    cbReceiveGoodsInAlertsSuperUser.Checked = False
                End If
            Else
                cbReceiveGoodsInAlertsProductOwner.Checked = False
                cbReceiveGoodsInAlertsSuperUser.Checked = False
            End If
            If Not IsDBNull(oDataReader("LowStockAlert")) Then
                If oDataReader("LowStockAlert") Then
                    cbReceiveLowStockAlertsProductOwner.Checked = True
                    cbReceiveLowStockAlertsSuperUser.Checked = True
                Else
                    cbReceiveLowStockAlertsProductOwner.Checked = False
                    cbReceiveLowStockAlertsSuperUser.Checked = False
                End If
            Else
                cbReceiveLowStockAlertsProductOwner.Checked = False
                cbReceiveLowStockAlertsSuperUser.Checked = False
            End If
            If Not IsDBNull(oDataReader("MemorableAnswer1")) Then
                tbMemorableAnswer1.Text = oDataReader("MemorableAnswer1")
            End If
            If Not IsDBNull(oDataReader("MemorableAnswer2")) Then
                tbMemorableAnswer2.Text = oDataReader("MemorableAnswer2")
            End If
            If Not IsDBNull(oDataReader("MemorableAnswer3")) Then
                tbMemorableAnswer3.Text = oDataReader("MemorableAnswer3")
            End If
            If Not IsDBNull(oDataReader("DefaultDestinationGABKey")) Then
                nDefaultDestinationGABKey = oDataReader("DefaultDestinationGABKey")
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try

        If nDefaultDestinationGABKey > 0 Then
            lblDefaultDestination.Text = GetDefaultDestination(nDefaultDestinationGABKey)
            If lblDefaultDestination.Text.Length > 0 Then
                trDefaultDestination.Visible = True
            End If
        Else
            trDefaultDestination.Visible = False
        End If
    End Sub

    Protected Function GetDefaultDestination(ByVal nAddressKey As Integer) As String
        GetDefaultDestination = String.Empty
        Dim sAddress As String = String.Empty
        Dim sTemp As String
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Address_GetAddressFromKey2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As New SqlParameter("@AddressKey", SqlDbType.Int, 4)
        oCmd.Parameters.Add(oParam)
        oParam.Value = nAddressKey
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If oDataReader("DeletedFlag") & String.Empty = "Y" Then
                Exit Function
            End If
            sAddress = oDataReader("AttnOf") & String.Empty
            If sAddress.Length > 0 Then
                sAddress += ", "
            End If
            sAddress += oDataReader("Company") & String.Empty & ", "
            sAddress += oDataReader("Addr1") & String.Empty & ", "
            sTemp = oDataReader("Addr2") & String.Empty
            If sTemp.Length > 0 Then
                sAddress += sTemp & ", "
            End If
            sAddress += oDataReader("Town") & String.Empty & ", "
            sAddress += oDataReader("PostCode") & String.Empty & ", "
            sAddress += oDataReader("CountryName") & String.Empty
            oDataReader.Close()
        Catch ex As Exception
            WebMsgBox.Show("GetDefaultDestination" & ex.ToString)
        End Try
        oConn.Close()
        GetDefaultDestination = sAddress
    End Function
    
    Protected Sub PopulateCMItemsDropdown()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_CMGetByCustomer", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)

        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        Dim arr As ArrayList = New ArrayList
        ddlCMItems.Items.Add(New ListItem("- please select -", 0))
        For Each row As Object In oDataReader
            Dim sValueDate As String = oDataReader("ProductDate") & String.Empty
            If sValueDate.Length > 0 Then
                sValueDate = " - " & sValueDate
            End If
            Dim sDescription As String = oDataReader("ProductDescription") & String.Empty
            If sDescription.Length > 30 Then
                sDescription = sDescription.Substring(0, 27).PadRight(30, ".")
            End If
            ddlCMItems.Items.Add(New ListItem(oDataReader("ProductCode") & sValueDate & " - " & sDescription, oDataReader("LogisticProductKey")))
        Next
        oConn.Close()
    End Sub

    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SaveChanges()
    End Sub
    
    Protected Function SaveLovellsPublicationUserChanges() As Integer
        Dim nReceivePublicationOrderAlerts As Integer = 0
        Dim nReceiveProductInactivityAlerts As Integer = 0
        If cbReceivePublicationOrderAlerts.Checked Then
            nReceivePublicationOrderAlerts = 1
        End If
        If cbReceiveProductInactivityAlerts.Checked Then
            nReceiveProductInactivityAlerts = 1
        End If
        If ExecuteQueryToDataTable("SELECT * FROM UserPublicationProfile WHERE UserKey = " & Session("UserKey")).Rows.Count > 0 Then
            SaveLovellsPublicationUserChanges = ExecuteQueryToDataTable("UPDATE UserPublicationProfile SET ReceivePublicationOrderAlerts = " & nReceivePublicationOrderAlerts & ", ReceiveProductInactivityAlerts = " & nReceiveProductInactivityAlerts & " WHERE UserKey = " & Session("UserKey") & " SELECT @@ROWCOUNT").Rows(0).Item(0)
        Else
            SaveLovellsPublicationUserChanges = ExecuteQueryToDataTable("INSERT INTO UserPublicationProfile (UserKey, ReceivePublicationOrderAlerts, ReceiveProductInactivityAlerts ) VALUES (" & Session("UserKey") & ", " & nReceivePublicationOrderAlerts & ", " & nReceiveProductInactivityAlerts & ") SELECT @@ROWCOUNT").Rows(0).Item(0)
        End If
    End Function
    
    Protected Sub SaveChanges()
        Dim bStockArrivalAlert As Boolean = False
        Dim bLowStockAlert As Boolean = False
        Dim bStockBookingAlert As Boolean = False
        Dim bStockBookingAlertAll As Boolean = False

        If Session("UserType").ToString.ToLower = "superuser" Then
            bLowStockAlert = cbReceiveLowStockAlertsSuperUser.Checked
            bStockArrivalAlert = cbReceiveGoodsInAlertsSuperUser.Checked
            bStockBookingAlertAll = cbReceiveOrderConfirmationAlertsSuperUser.Checked
        End If

        If Session("UserType").ToString.ToLower.Contains("owner") Then
            bLowStockAlert = cbReceiveLowStockAlertsProductOwner.Checked
            bStockArrivalAlert = cbReceiveGoodsInAlertsProductOwner.Checked
            bStockBookingAlertAll = cbReceiveOrderConfirmationAlertsProductOwner.Checked
        End If
        
        If IsLovells() And Session("UserType").ToString.ToLower = "user" Then
            Call SaveLovellsPublicationUserChanges()
        End If

        bStockBookingAlert = cbReceiveOrderConfirmationAlertsUser.Checked
        
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_UpdateMyProfile3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = CLng(Session("UserKey"))
        oCmd.Parameters.Add(paramUserKey)

        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        paramCustomerKey.Value = CLng(Session("CustomerKey"))
        oCmd.Parameters.Add(paramCustomerKey)

        Dim paramDepartment As SqlParameter = New SqlParameter("@Department", SqlDbType.NVarChar, 20)
        paramDepartment.Value = tbCostCentre.Text
        oCmd.Parameters.Add(paramDepartment)

        Dim paramEmailAddr As SqlParameter = New SqlParameter("@EmailAddr", SqlDbType.NVarChar, 100)
        paramEmailAddr.Value = tbEmail.Text
        oCmd.Parameters.Add(paramEmailAddr)

        Dim paramTelephone As SqlParameter = New SqlParameter("@Telephone", SqlDbType.NVarChar, 20)
        paramTelephone.Value = tbTelephone.Text
        oCmd.Parameters.Add(paramTelephone)
        
        Dim paramStockBookingAlert As SqlParameter = New SqlParameter("@StockBookingAlert", SqlDbType.Bit, 1)
        paramStockBookingAlert.Value = bStockBookingAlert
        oCmd.Parameters.Add(paramStockBookingAlert)
        
        Dim paramStockBookingAlertAll As SqlParameter = New SqlParameter("@StockBookingAlertAll", SqlDbType.Bit, 1)
        paramStockBookingAlertAll.Value = bStockBookingAlertAll
        oCmd.Parameters.Add(paramStockBookingAlertAll)
        
        Dim paramStockArrivalAlert As SqlParameter = New SqlParameter("@StockArrivalAlert", SqlDbType.Bit, 1)
        paramStockArrivalAlert.Value = bStockArrivalAlert
        oCmd.Parameters.Add(paramStockArrivalAlert)
        
        Dim paramLowStockAlert As SqlParameter = New SqlParameter("@LowStockAlert", SqlDbType.Bit, 1)
        paramLowStockAlert.Value = bLowStockAlert
        oCmd.Parameters.Add(paramLowStockAlert)

        Dim paramMemorableAnswer1 As SqlParameter = New SqlParameter("@MemorableAnswer1", SqlDbType.NVarChar, 50)
        paramMemorableAnswer1.Value = tbMemorableAnswer1.Text
        oCmd.Parameters.Add(paramMemorableAnswer1)
        
        Dim paramMemorableAnswer2 As SqlParameter = New SqlParameter("@MemorableAnswer2", SqlDbType.NVarChar, 50)
        paramMemorableAnswer2.Value = tbMemorableAnswer2.Text
        oCmd.Parameters.Add(paramMemorableAnswer2)
        
        Dim paramMemorableAnswer3 As SqlParameter = New SqlParameter("@MemorableAnswer3", SqlDbType.NVarChar, 50)
        paramMemorableAnswer3.Value = tbMemorableAnswer3.Text
        oCmd.Parameters.Add(paramMemorableAnswer3)
        
        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try
        WebMsgBox.Show("Your profile has been updated")
    End Sub
    
    Protected Sub GetProductCodeReservations()
        Dim sSQL As String = "SELECT * FROM ProductCodeReservation WHERE UserKey = " & Session("UserKey")
        Dim oConn As New SqlConnection(gsConn)
        Dim dtReservations As New DataTable
        Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
        oAdapter.Fill(dtReservations)
        If dtReservations.Rows.Count > 0 Then
            gvProductCodeReservations.Visible = True
            gvProductCodeReservations.DataSource = dtReservations
            gvProductCodeReservations.DataBind()
            lblReservationMessage.Text = "Product Code Reservations"
            lblReservationMessage.Font.Bold = True
        Else
            gvProductCodeReservations.Visible = False
            lblReservationMessage.Text = "You have no product code reservations"
            lblReservationMessage.Font.Bold = False
        End If
        oConn.Close()
    End Sub
    
    Protected Sub lnkbtnCancelReservation_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call DeleteReservation(lb.CommandArgument)
        Call GetProductCodeReservations()
    End Sub
    
    Protected Sub DeleteReservation(ByVal id As Long)
        Dim sSQL As String = "DELETE FROM ProductCodeReservation WHERE [id] = " & id
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        oConn.Open()
        oCmd = New SqlCommand(sSQL, oConn)
        oCmd.ExecuteNonQuery()
        oConn.Close()
    End Sub

    Protected Sub GetPendingOrderAuthorisations()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_AuthOrderGetPendingForOrderer", oConn)
        
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserProfileKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@UserProfileKey").Value = Session("UserKey")

        oAdapter.Fill(oDatatable)
        gvAuthoriseOrder.DataSource = oDatatable
        gvAuthoriseOrder.DataBind()
    End Sub
    
    Protected Sub btnDeleteOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnDeleteOrder As Button = sender
        Call UpdateHoldingQueueEntry(sStatus:="DELETED", lConsignmentKey:=0, nHoldingQueueKey:=btnDeleteOrder.CommandArgument)
        Call GetPendingOrderAuthorisations()
    End Sub

    Protected Sub btnViewOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnViewOrder As Button = sender
        GetAuthorisationOrder(CInt(btnViewOrder.CommandArgument))
        hidHoldingQueueKey.Value = btnViewOrder.CommandArgument
        Call ShowOrderDetailPanel()
    End Sub
    
    Protected Sub GetAuthorisationOrder(ByVal nHoldingQueueKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As DataTable = GetAuthOrderDetails(nHoldingQueueKey)
        Dim drOrderDetails As DataRow = GetOrderAuthorisationByKey(nHoldingQueueKey)
        lblAuthOrderOrderedBy.Text = drOrderDetails.Item("FirstName") & " " & drOrderDetails.Item("LastName")
        lblAuthOrderPlacedOn.Text = drOrderDetails.Item("OrderCreatedDateTime")
        lblAuthOrderConsignee.Text = drOrderDetails.Item("CneeName")
        lblAuthOrderAttnOf.Text = drOrderDetails.Item("CneeCtcName")
        lblAuthOrderAddr1.Text = drOrderDetails.Item("CneeAddr1")
        lblAuthOrderAddr2.Text = drOrderDetails.Item("CneeAddr2")
        lblAuthOrderAddr3.Text = drOrderDetails.Item("CneeAddr3")
        lblAuthOrderTown.Text = drOrderDetails.Item("CneeTown")
        lblAuthOrderState.Text = drOrderDetails.Item("CneeState")
        lblAuthOrderPostcode.Text = drOrderDetails.Item("CneePostCode")
        lblAuthOrderCountry.Text = drOrderDetails.Item("CountryName")
        gvAuthOrderDetails.DataSource = oDatatable
        gvAuthOrderDetails.DataBind()
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
    
    Protected Sub UpdateHoldingQueueEntry(ByVal sStatus As String, ByVal lConsignmentKey As Long, ByVal nHoldingQueueKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_StockBooking_AuthOrderUpdateHoldingQueue", oConn)
        Dim spParam As SqlParameter
        oCmd.CommandType = CommandType.StoredProcedure

        spParam = New SqlParameter("@HoldingQueueKey", SqlDbType.Int)
        spParam.Value = nHoldingQueueKey
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@OrderStatus", SqlDbType.NVarChar, 50)
        spParam.Value = sStatus
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@ConsignmentKey", SqlDbType.Int)
        spParam.Value = lConsignmentKey
        oCmd.Parameters.Add(spParam)

        spParam = New SqlParameter("@MsgToOrderer", SqlDbType.NVarChar, 1000)
        spParam.Value = String.Empty
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
    
    Protected Sub btnGoBack_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowAuthorisationsPanel()
    End Sub

    Protected Sub btnDeleteOrderFromDetailScreen_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call UpdateHoldingQueueEntry(sStatus:="DELETED", lConsignmentKey:=0, nHoldingQueueKey:=hidHoldingQueueKey.Value)
        Call GetPendingOrderAuthorisations()
        Call ShowAuthorisationsPanel()
    End Sub
    
    Protected Function gvAuthOrderDetailsItemForeColor(ByVal DataItem As Object) As System.Drawing.Color
        gvAuthOrderDetailsItemForeColor = Black
        If Not IsDBNull(DataBinder.Eval(DataItem, "Authorised")) AndAlso DataBinder.Eval(DataItem, "Authorised") = "N" Then
            gvAuthOrderDetailsItemForeColor = Red
        End If
    End Function

    Protected Sub btnCMShowEvents_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sCommandType As String = b.CommandArgument  ' my or all
        Call ShowEvents(bEventTypeAll:=b.CommandArgument.ToLower = "all", bByProduct:=rbShowByItem.Checked)
    End Sub
    
    Protected Sub gvCalendarManagedItems_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Call RebindEventsGrid()
    End Sub
    
    Protected Sub RebindEventsGrid()
        Select Case pnDisplayMode
            Case 1, 2
                Call ShowEvents(bEventTypeAll:=True, bByProduct:=rbShowByItem.Checked)
            Case 3, 4
                Call ShowEvents(bEventTypeAll:=False, bByProduct:=rbShowByItem.Checked)
        End Select
    End Sub
    
    Protected Sub ShowEvents(ByVal bEventTypeAll As Boolean, ByVal bByProduct As Boolean)   ' false=my events / true=all events, false=by event / true = by product
        If bEventTypeAll Then
            If bByProduct Then
                Call ShowAllEventsByProduct()
                lblLegendCalendarManagedItems.Text = "Calendar Managed Items - All Events By Product"
                pnDisplayMode = 1
            Else
                Call ShowAllEvents()
                lblLegendCalendarManagedItems.Text = "Calendar Managed Items - All Events"
                pnDisplayMode = 2
            End If
        Else
            If bByProduct Then
                Call ShowMyEventsByProduct()
                lblLegendCalendarManagedItems.Text = "Calendar Managed Items - My Events By Product"
                pnDisplayMode = 3
            Else
                Call ShowMyEvents()
                lblLegendCalendarManagedItems.Text = "Calendar Managed Items - My Events"
                pnDisplayMode = 4
            End If
        End If
    End Sub
    
    Protected Sub ShowMyEvents()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL1 As New StringBuilder
        Dim bFirstTime As Boolean = True
        sbSQL1.Append("SELECT DISTINCT EventId INTO #EventId FROM CalendarManagedItemDays cmid INNER JOIN CalendarManagedItemEvent cmie ON cmid.EventId = cmie.id WHERE EventDay >= GETDATE() AND BookedBy = " & Session("UserKey"))
        Dim oCmd1 As SqlCommand = New SqlCommand(sbSQL1.ToString, oConn)
        oConn.Open()
        oCmd1.ExecuteNonQuery()
        Dim sbSQL2 As New StringBuilder
        sbSQL2.Append("SELECT ei.EventId, BookedBy, EventName Event, FirstName + ' ' + LastName 'Booked by', CONVERT(VARCHAR(9), MIN(EventDay), 6) 'Delivery Date', CONVERT(VARCHAR(9), MAX(EventDay), 6) 'Collection Date' ")
        sbSQL2.Append("FROM #EventId ei INNER JOIN CalendarManagedItemDays cmid ON ei.EventId = cmid.EventId INNER JOIN CalendarManagedItemEvent cmie ON ei.EventId = cmie.id INNER JOIN UserProfile up ON cmie.BookedBy = up.[Key] ")
        sbSQL2.Append("WHERE cmie.IsDeleted = 0 OR cmie.IsDeleted IS NULL ")
        sbSQL2.Append("GROUP BY ei.EventId, EventName, FirstName, LastName, BookedBy ")
        sbSQL2.Append("ORDER BY  MIN(EventDay)")
        Dim oCmd2 As SqlCommand = New SqlCommand(sbSQL2.ToString, oConn)
        oDataReader = oCmd2.ExecuteReader()
        Dim arr As ArrayList = New ArrayList
        For Each row As Object In oDataReader
            arr.Add(row)
        Next
        gvCalendarManagedItems.DataSource = arr
        gvCalendarManagedItems.DataBind()
        oConn.Close()
        gvCalendarManagedItems.Visible = True
    End Sub
    
    Protected Sub ShowAllEvents()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim sbSQL1 As New StringBuilder
        Dim bFirstTime As Boolean = True
        sbSQL1.Append("SELECT DISTINCT EventId INTO #EventId FROM CalendarManagedItemDays cmid INNER JOIN CalendarManagedItemEvent cmie ON cmid.EventId = cmie.id WHERE EventDay >= GETDATE() AND CustomerKey = " & Session("CustomerKey"))
        Dim oCmd1 As SqlCommand = New SqlCommand(sbSQL1.ToString, oConn)
        oConn.Open()
        oCmd1.ExecuteNonQuery()
        Dim sbSQL2 As New StringBuilder
        sbSQL2.Append("SELECT ei.EventId, BookedBy, EventName Event, FirstName + ' ' + LastName 'Booked by', CONVERT(VARCHAR(9), MIN(EventDay), 6) 'Delivery Date', CONVERT(VARCHAR(9), MAX(EventDay), 6) 'Collection Date' ")
        sbSQL2.Append("FROM #EventId ei INNER JOIN CalendarManagedItemDays cmid ON ei.EventId = cmid.EventId INNER JOIN CalendarManagedItemEvent cmie ON ei.EventId = cmie.id INNER JOIN UserProfile up ON cmie.BookedBy = up.[Key] ")
        sbSQL2.Append("WHERE cmie.IsDeleted = 0 OR cmie.IsDeleted IS NULL ")
        sbSQL2.Append("GROUP BY ei.EventId, EventName, FirstName, LastName, BookedBy ")
        sbSQL2.Append("ORDER BY  MIN(EventDay)")
        Dim oCmd2 As SqlCommand = New SqlCommand(sbSQL2.ToString, oConn)
        oDataReader = oCmd2.ExecuteReader()
        Dim arr As ArrayList = New ArrayList
        For Each row As Object In oDataReader
            arr.Add(row)
        Next
        gvCalendarManagedItems.DataSource = arr
        gvCalendarManagedItems.DataBind()
        oConn.Close()
        gvCalendarManagedItems.Visible = True
    End Sub
    
    Protected Sub ShowMyEventsByProduct()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_GetEventsByProductByUser2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamLogisticProductKey As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int)
        oParamLogisticProductKey.Value = ddlCMItems.SelectedValue
        Dim oParamUserKey As SqlParameter = oCmd.Parameters.Add("@UserKey", SqlDbType.Int)
        oParamUserKey.Value = Session("UserKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            Dim arr As ArrayList = New ArrayList
            For Each row As Object In oDataReader
                arr.Add(row)
            Next
            gvCalendarManagedItems.DataSource = arr
            gvCalendarManagedItems.DataBind()
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve my events by product; " & ex.Message)
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
        gvCalendarManagedItems.Visible = True
    End Sub
    
    Protected Sub ShowAllEventsByProduct()
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_GetEventsByProduct2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamLogisticProductKey As SqlParameter = oCmd.Parameters.Add("@LogisticProductKey", SqlDbType.Int)
        oParamLogisticProductKey.Value = ddlCMItems.SelectedValue
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            Dim arr As ArrayList = New ArrayList
            For Each row As Object In oDataReader
                arr.Add(row)
            Next
            gvCalendarManagedItems.DataSource = arr
            gvCalendarManagedItems.DataBind()
        Catch ex As SqlException
            WebMsgBox.Show("Internal error - could not retrieve events by product; " & ex.Message)
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
        gvCalendarManagedItems.Visible = True
    End Sub
    
    Protected Sub ddlCMItems_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = 0 Then
            btnCMShowMyEvents.Enabled = False
            btnCMShowAllEvents.Enabled = False
        Else
            btnCMShowMyEvents.Enabled = True
            btnCMShowAllEvents.Enabled = True
        End If
        gvCalendarManagedItems.Visible = False
    End Sub
    
    Protected Sub rbShowByItem_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            If ddlCMItems.Items.Count = 0 Then
                Call PopulateCMItemsDropdown()
            End If
            btnCMShowMyEvents.Enabled = False
            btnCMShowAllEvents.Enabled = False
            lblCMLegendSelectItem.Visible = True
            ddlCMItems.Visible = True
            gvCalendarManagedItems.Visible = False
            lblLegendCalendarManagedItems.Text = "Calendar Managed Items"
        End If
    End Sub
    
    Protected Sub rbShowByEvent_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            lblCMLegendSelectItem.Visible = False
            ddlCMItems.Visible = False
            ddlCMItems.SelectedIndex = 0
            btnCMShowMyEvents.Enabled = True
            btnCMShowAllEvents.Enabled = True
            gvCalendarManagedItems.Visible = False
            lblLegendCalendarManagedItems.Text = "Calendar Managed Items"
        End If
    End Sub
    
    Protected Sub gvCalendarManagedItems_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.RowType = DataControlRowType.DataRow Then
            
        End If
        If row.Cells.Count >= 3 Then      ' check if one or more rows - if no rows there will only be a single cell with the empty grid message
            row.Cells(1).Visible = False  ' hide items required in the query but not to be displayed (EventId, BookedBy)
            row.Cells(2).Visible = False
        End If
    End Sub

    Protected Function gvCMSetButtonVisibility(ByVal DataItem As Object) As String
        If DataBinder.Eval(DataItem, "BookedBy") = Session("UserKey") Then
            Return True
        Else
            Return False
        End If
    End Function
    
    Protected Sub lnkbtnCMShowDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        pnEventId = CInt(lb.CommandArgument)
        Call CMShowEvent(lb.CommandArgument, lb.CommandName)
    End Sub

    Protected Sub InitCountryDropdowns()
        If ddlCMCountry.Items.Count = 0 Or ddlCMCollectionCountry.Items.Count = 0 Then
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
        End If
    End Sub
    
    Protected Sub CMShowEvent(ByVal nEventId As Integer, ByVal bSaveEventChangesVisibility As Boolean)
        Call InitCountryDropdowns()
        Call GetEventFromId()
        Call GetNotes()
        Call ShowEventPanel()
        Dim dtDeliveryDate As Date = Date.Parse(lblDeliveryDate.Text)
        lblOnlineChangesMessage.Visible = False
        If bSaveEventChangesVisibility Then
            If DateDiff(DateInterval.Day, Date.Now, dtDeliveryDate) <= 5 Then
                bSaveEventChangesVisibility = False
                lblOnlineChangesMessage.Visible = True
            End If
        End If
        btnSaveEventChanges.Visible = bSaveEventChangesVisibility
    End Sub
    
    Protected Sub GetEventFromId()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable1 As New DataTable
        Dim oAdapter1 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventById", oConn)
        Dim nDDLIndex As Integer

        oAdapter1.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter1.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter1.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        Try
            oConn.Open()
            oAdapter1.Fill(oDataTable1)
            Dim dr As DataRow = oDataTable1.Rows(0)
            ' need to check 1 and only 1 row present
            lblCustomer.Text = dr("CustomerAccountCode")
            lblEventName.Text = dr("EventName")
            tbContactName.Text = dr("ContactName")
            tbContactPhone.Text = dr("ContactPhone")
            tbContactMobile.Text = dr("ContactMobile")
            
            Dim sContactName2 As String
            Dim sContactPhone2 As String
            Dim sContactMobile2 As String
            
            If Not IsDBNull(dr("ContactName2")) Then
                sContactName2 = dr("ContactName2").ToString.Trim
            Else
                sContactName2 = String.Empty
            End If
            If Not IsDBNull(dr("ContactPhone2")) Then
                sContactPhone2 = dr("ContactPhone2").ToString.Trim
            Else
                sContactPhone2 = String.Empty
            End If
            If Not IsDBNull(dr("ContactMobile2")) Then
                sContactMobile2 = dr("ContactMobile2").ToString.Trim
            Else
                sContactMobile2 = String.Empty
            End If
            
            tbCMContactName2.Text = sContactName2
            tbCMContactPhone2.Text = sContactPhone2
            tbCMContactMobile2.Text = sContactMobile2
            If String.IsNullOrEmpty(sContactName2) And String.IsNullOrEmpty(sContactPhone2) And String.IsNullOrEmpty(sContactMobile2) Then
                Call SetContact2FieldsVisibility(False)
            Else
                Call SetContact2FieldsVisibility(True)
            End If
            
            tbEventAddress1.Text = dr("EventAddress1")
            tbEventAddress2.Text = dr("EventAddress2")
            tbEventAddress3.Text = dr("EventAddress3")
            tbTown.Text = dr("Town")
            tbPostcode.Text = dr("Postcode")
            
            Dim nCountryKey As Integer
            If Not IsDBNull(dr("CountryKey")) Then
                nCountryKey = dr("CountryKey")
            Else
                nCountryKey = COUNTRY_KEY_UK
            End If

            If nCountryKey = COUNTRY_KEY_UK Then
                trCMCountry.Visible = False
                lnkbtnCMAddressOutsideUK.Visible = True
            Else
                trCMCountry.Visible = True
                lnkbtnCMAddressOutsideUK.Visible = False
            End If

            For nDDLIndex = 1 To ddlCMCountry.Items.Count - 1
                If ddlCMCountry.Items(nDDLIndex).Value = nCountryKey Then
                    ddlCMCountry.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            
            Dim sTemp As String = dr("DeliveryTime")
            For nDDLIndex = 0 To ddlDeliveryTime.Items.Count - 1
                If ddlDeliveryTime.Items(nDDLIndex).Text = sTemp Then
                    ddlDeliveryTime.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            tbPreciseDeliveryPoint.Text = dr("PreciseDeliveryPoint")
            tbPreciseCollectionPoint.Text = dr("PreciseCollectionPoint")
            sTemp = dr("CollectionTime")
            For nDDLIndex = 0 To ddlCollectionTime.Items.Count - 1
                If ddlCollectionTime.Items(nDDLIndex).Text = sTemp Then
                    ddlCollectionTime.SelectedIndex = nDDLIndex
                    Exit For
                End If
            Next
            If Not IsDBNull(dr("DifferentCollectionAddress")) Then
                cbDifferentCollectionAddress.Checked = dr("DifferentCollectionAddress")
            Else
                cbDifferentCollectionAddress.Checked = False
            End If
            SetCollectionAddressVisibility(cbDifferentCollectionAddress.Checked)
            If cbDifferentCollectionAddress.Checked Then
                tbCollectionAddress1.Text = dr("CollectionAddress1")
                tbCollectionAddress2.Text = dr("CollectionAddress2")
                tbCollectionTown.Text = dr("CollectionTown")
                tbCollectionPostcode.Text = dr("CollectionPostcode")

                Dim nCollectionCountryKey As Integer
                If Not IsDBNull(dr("CountryKey")) Then
                    nCollectionCountryKey = dr("CountryKey")
                Else
                    nCollectionCountryKey = COUNTRY_KEY_UK
                End If

                If nCollectionCountryKey = COUNTRY_KEY_UK Then
                    trCMCollectionCountry.Visible = False
                    lnkbtnCMCollectionAddressOutsideUK.Visible = True
                Else
                    trCMCollectionCountry.Visible = True
                    lnkbtnCMCollectionAddressOutsideUK.Visible = False
                End If

                For nDDLIndex = 1 To ddlCMCollectionCountry.Items.Count - 1
                    If ddlCMCollectionCountry.Items(nDDLIndex).Value = nCollectionCountryKey Then
                        ddlCMCollectionCountry.SelectedIndex = nDDLIndex
                        Exit For
                    End If
                Next
            
            End If
            If Not IsDBNull(dr("CustomerReference")) Then
                tbCustomerReference.Text = dr("CustomerReference")
            Else
                tbCustomerReference.Text = String.Empty
            End If
            tbSpecialInstructions.Text = dr("SpecialInstructions")
            lblBookedBy.Text = dr("username")
            lblBookedOn.Text = dr("BookedOn")
            
            Dim oAdapter2 As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventItemsById", oConn)
            
            oAdapter2.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter2.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
            oAdapter2.SelectCommand.Parameters("@EventId").Value = pnEventId
            Dim oDataTable2 As New DataTable
            oAdapter2.Fill(oDataTable2)
            gvItems.DataSource = oDataTable2
            gvItems.DataBind()
            
            If gvItems.Rows.Count = 1 Then
                lblLegendProduct.Text = "Product:"
            Else
                lblLegendProduct.Text = "Products:"
            End If
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try

        For Each gvr As GridViewRow In gvCalendarManagedItems.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim lb As LinkButton = gvr.FindControl("lnkbtnCMShowDetails")
                If lb.CommandArgument = pnEventId Then
                    lblDeliveryDate.Text = gvr.Cells(5).Text
                    lblCollectionDate.Text = gvr.Cells(6).Text
                End If
            End If
        Next
    End Sub
    
    Protected Sub SetContact2FieldsVisibility(ByVal bVisibility As Boolean)
        lnkbtnCMAddSecondContact.Visible = Not bVisibility
        rfvCMContactName2.Visible = bVisibility
        lblLegendCMContactName2.Visible = bVisibility
        tbCMContactName2.Visible = bVisibility
        rfvCMContactMobile2.Visible = bVisibility
        lblLegendCMContactMobile2.Visible = bVisibility
        tbCMContactMobile2.Visible = bVisibility
        trCMContactPhone2.Visible = bVisibility
    End Sub
    
    Protected Sub lnkbtnCMDeleteEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        Call CMDeleteEvent(lb.CommandArgument)
        pnlEvent.Visible = False
        Call RebindEventsGrid()
    End Sub

    Protected Sub CMDeleteEvent(ByVal nEventId As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_DeleteEvent", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = nEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramUserId As SqlParameter = New SqlParameter("@UserId", SqlDbType.Int)
        paramUserId.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserId)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub gvNotes_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Dim gv As GridView = sender
        gv.PageIndex = e.NewPageIndex
        Call GetNotes()
    End Sub

    Protected Sub GetNotes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_CalendarManaged_GetEventNotes2", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@EventId", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@EventId").Value = pnEventId
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerVisibleOnly", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@CustomerVisibleOnly").Value = 1
        
        Try
            oConn.Open()
            oAdapter.Fill(oDataTable)
            gvNotes.DataSource = oDataTable
            gvNotes.DataBind()
        Catch ex As Exception
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnShowHideNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ToggleNotesGrid()
    End Sub
    
    Protected Sub lnkbtnRefreshNotes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GetNotes()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("show") Then
            Call ToggleNotesGrid()
        End If
    End Sub
    
    Protected Sub ToggleNotesGrid()
        If lnkbtnShowHideNotes.Text.ToLower.Contains("hide") Then
            lnkbtnShowHideNotes.Text = "show notes"
            trNotes.Visible = False
        Else
            lnkbtnShowHideNotes.Text = "hide notes"
            trNotes.Visible = True
        End If
    End Sub
    
    Protected Sub btnSaveEventChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate("CalendarManaged")
        If Page.IsValid Then
            Call SaveEventChanges()
            WebMsgBox.Show("Your changes were saved.")
            Dim bOverseasBooking As Boolean = False
            If ddlCMCountry.Visible Then
                If ddlCMCountry.SelectedValue <> COUNTRY_KEY_UK Then
                    bOverseasBooking = True
                End If
            End If
            If ddlCMCollectionCountry.Visible Then
                If ddlCMCollectionCountry.SelectedValue <> COUNTRY_KEY_UK Then
                    bOverseasBooking = True
                End If
            End If
            If bOverseasBooking Then
                Call AlertOverseasBooking()
            End If
        Else
            WebMsgBox.Show("One or more fields were incorrect or not supplied. Please correct the information and resubmit.")
        End If
    End Sub
    
    Protected Sub SaveEventChanges()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_CalendarManaged_UpdateEvent3", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramEventId As SqlParameter = New SqlParameter("@EventId", SqlDbType.Int)
        paramEventId.Value = pnEventId
        oCmd.Parameters.Add(paramEventId)

        Dim paramContactName As SqlParameter = New SqlParameter("@ContactName", SqlDbType.VarChar, 50)
        paramContactName.Value = tbContactName.Text
        oCmd.Parameters.Add(paramContactName)

        Dim paramContactPhone As SqlParameter = New SqlParameter("@ContactPhone", SqlDbType.VarChar, 50)
        paramContactPhone.Value = tbContactPhone.Text
        oCmd.Parameters.Add(paramContactPhone)

        Dim paramContactMobile As SqlParameter = New SqlParameter("@ContactMobile", SqlDbType.VarChar, 50)
        paramContactMobile.Value = tbContactMobile.Text
        oCmd.Parameters.Add(paramContactMobile)

        Dim paramContactName2 As SqlParameter = New SqlParameter("@ContactName2", SqlDbType.VarChar, 50)
        paramContactName2.Value = tbCMContactName2.Text
        oCmd.Parameters.Add(paramContactName2)

        Dim paramContactPhone2 As SqlParameter = New SqlParameter("@ContactPhone2", SqlDbType.VarChar, 50)
        paramContactPhone2.Value = tbCMContactPhone2.Text
        oCmd.Parameters.Add(paramContactPhone2)

        Dim paramContactMobile2 As SqlParameter = New SqlParameter("@ContactMobile2", SqlDbType.VarChar, 50)
        paramContactMobile2.Value = tbCMContactMobile2.Text
        oCmd.Parameters.Add(paramContactMobile2)

        Dim paramEventAddress1 As SqlParameter = New SqlParameter("@EventAddress1", SqlDbType.VarChar, 50)
        paramEventAddress1.Value = tbEventAddress1.Text
        oCmd.Parameters.Add(paramEventAddress1)

        Dim paramEventAddress2 As SqlParameter = New SqlParameter("@EventAddress2", SqlDbType.VarChar, 50)
        paramEventAddress2.Value = tbEventAddress2.Text
        oCmd.Parameters.Add(paramEventAddress2)

        Dim paramEventAddress3 As SqlParameter = New SqlParameter("@EventAddress3", SqlDbType.VarChar, 50)
        paramEventAddress3.Value = tbEventAddress3.Text
        oCmd.Parameters.Add(paramEventAddress3)

        Dim paramTown As SqlParameter = New SqlParameter("@Town", SqlDbType.VarChar, 50)
        paramTown.Value = tbTown.Text
        oCmd.Parameters.Add(paramTown)

        Dim paramPostcode As SqlParameter = New SqlParameter("@Postcode", SqlDbType.VarChar, 50)
        paramPostcode.Value = tbPostcode.Text
        oCmd.Parameters.Add(paramPostcode)

        Dim paramCountryKey As SqlParameter = New SqlParameter("@CountryKey", SqlDbType.Int)
        paramCountryKey.Value = ddlCMCountry.SelectedValue
        oCmd.Parameters.Add(paramCountryKey)

        Dim paramDeliveryTime As SqlParameter = New SqlParameter("@DeliveryTime", SqlDbType.VarChar, 50)
        paramDeliveryTime.Value = ddlDeliveryTime.SelectedItem.Text
        oCmd.Parameters.Add(paramDeliveryTime)

        Dim paramPreciseDeliveryPoint As SqlParameter = New SqlParameter("@PreciseDeliveryPoint", SqlDbType.VarChar, 100)
        paramPreciseDeliveryPoint.Value = tbPreciseDeliveryPoint.Text
        oCmd.Parameters.Add(paramPreciseDeliveryPoint)

        Dim paramDifferentCollectionAddress As SqlParameter = New SqlParameter("@DifferentCollectionAddress", SqlDbType.Bit)
        paramDifferentCollectionAddress.Value = cbDifferentCollectionAddress.Checked
        oCmd.Parameters.Add(paramDifferentCollectionAddress)

        Dim paramCollectionAddress1 As SqlParameter = New SqlParameter("@CollectionAddress1", SqlDbType.NVarChar, 50)
        paramCollectionAddress1.Value = tbCollectionAddress1.Text
        oCmd.Parameters.Add(paramCollectionAddress1)

        Dim paramCollectionAddress2 As SqlParameter = New SqlParameter("@CollectionAddress2", SqlDbType.NVarChar, 50)
        paramCollectionAddress2.Value = tbCollectionAddress2.Text
        oCmd.Parameters.Add(paramCollectionAddress2)

        Dim paramCollectionTown As SqlParameter = New SqlParameter("@CollectionTown", SqlDbType.NVarChar, 50)
        paramCollectionTown.Value = tbCollectionTown.Text
        oCmd.Parameters.Add(paramCollectionTown)

        Dim paramCollectionPostcode As SqlParameter = New SqlParameter("@CollectionPostcode", SqlDbType.NVarChar, 50)
        paramCollectionPostcode.Value = tbCollectionPostcode.Text
        oCmd.Parameters.Add(paramCollectionPostcode)

        Dim paramCollectionCountryKey As SqlParameter = New SqlParameter("@CollectionCountryKey", SqlDbType.Int)
        paramCollectionCountryKey.Value = ddlCMCollectionCountry.SelectedValue
        oCmd.Parameters.Add(paramCollectionCountryKey)

        Dim paramCollectionTime As SqlParameter = New SqlParameter("@CollectionTime", SqlDbType.VarChar, 50)
        paramCollectionTime.Value = ddlCollectionTime.SelectedItem.Text
        oCmd.Parameters.Add(paramCollectionTime)

        Dim paramPreciseCollectionPoint As SqlParameter = New SqlParameter("@PreciseCollectionPoint", SqlDbType.VarChar, 100)
        paramPreciseCollectionPoint.Value = tbPreciseCollectionPoint.Text
        oCmd.Parameters.Add(paramPreciseCollectionPoint)

        Dim paramSpecialInstructions As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.VarChar, 200)
        paramSpecialInstructions.Value = tbSpecialInstructions.Text
        oCmd.Parameters.Add(paramSpecialInstructions)

        Dim paramCustomerReference As SqlParameter = New SqlParameter("@CustomerReference", SqlDbType.NVarChar, 100)
        paramCustomerReference.Value = tbCustomerReference.Text
        oCmd.Parameters.Add(paramCustomerReference)

        Dim paramUpdatedBy As SqlParameter = New SqlParameter("@UpdatedBy", SqlDbType.Int)
        paramUpdatedBy.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUpdatedBy)

        Try
            oConn.Open()
            oCmd.Connection = oConn
            oCmd.ExecuteNonQuery()
        Catch ex As SqlException
            WebMsgBox.Show(ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub lnkbtnShowHideProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        If lb.Text.ToLower.Contains("show") Then
            lb.Text = "hide profile"
            pnlProfile.Visible = True
        Else
            lb.Text = "show profile"
            pnlProfile.Visible = False
        End If
    End Sub
    
    Protected Sub gvNotes_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.Cells.Count >= 3 Then
            row.Cells(3).Visible = False  ' hide Customer Visible flag
        End If
    End Sub
    
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

    Protected Sub lnkbtnCancelDefaultDestination_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call CancelDefaultDestination()
    End Sub
    
    Protected Sub CancelDefaultDestination()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_UserProfile_SetDefaultDestination", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim paramUserKey As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
        paramUserKey.Value = Session("UserKey")
        oCmd.Parameters.Add(paramUserKey)

        Dim paramAddressKey As SqlParameter = New SqlParameter("@DefaultDestinationGABKey", SqlDbType.Int, 4)
        paramAddressKey.Value = 0
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
        trDefaultDestination.Visible = False
    End Sub
    
    Protected Sub cbDifferentCollectionAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        Call SetCollectionAddressVisibility(cb.Checked)
    End Sub
    
    Protected Sub SetCollectionAddressVisibility(ByVal bVisibility As Boolean)
        trCollectionAddress1.Visible = bVisibility
        trCollectionAddress2.Visible = bVisibility
        If Not bVisibility Then
            tbCollectionAddress1.Text = String.Empty
            tbCollectionAddress2.Text = String.Empty
            tbCollectionTown.Text = String.Empty
            tbCollectionPostcode.Text = String.Empty
        End If
    End Sub
    
    Protected Sub lnkbtnCMAddSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetContact2FieldsVisibility(True)
        tbCMContactName2.Focus()
    End Sub

    Protected Sub lnkbtnCMAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCountry.Visible = True
        ddlCMCountry.SelectedIndex = 0
        ddlCMCountry.Focus()
    End Sub

    Protected Sub lnkbtnCMCollectionAddressOutsideUK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trCMCollectionCountry.Visible = True
        ddlCMCollectionCountry.SelectedIndex = 0
        ddlCMCollectionCountry.Focus()
    End Sub
    
    Protected Sub lnkbtnCMRemoveSecondContact_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbCMContactName2.Text = String.Empty
        tbCMContactPhone2.Text = String.Empty
        tbCMContactMobile2.Text = String.Empty
        Call SetContact2FieldsVisibility(False)
    End Sub
    
    Protected Sub AlertOverseasBooking()
        Dim sSQL As String
        Dim sRecipientName As String
        Dim sRecipientEmail As String
        Dim sText As String
        sSQL = "SELECT ISNULL(ah.AccountHandlerName,'Someone'), ISNULL(ah.EmailAddr,'') FROM AccountHandler ah INNER JOIN Customer c WHERE c.CustomerKey = " & Session("CustomerKey") & " AND ah.DeletedFlag <> 1"
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDataTable.Rows.Count > 0 Then
            Dim dr As DataRow = oDataTable.Rows(0)
            sRecipientName = dr(0)
            sRecipientEmail = dr(1).ToString.Trim
            If sRecipientEmail = String.Empty Then
                sRecipientEmail = "customer_services@sprintexpress.co.uk"
            End If
        Else
            sRecipientName = "Account Handler"
            sRecipientEmail = "customer_services@sprintexpress.co.uk"
        End If
        sText = "User " & Session("UserName") & ", Customer " & Session("Customer") & ", has just booked an Event outside the UK. The event name is " & lblEventName.Text
        Call SendMail("OVERSEAS EVENT ALERT", sRecipientEmail, "OVERSEAS EVENT SYSTEM ALERT", sText, sText)
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
            WebMsgBox.Show("Error in ExecuteNonQuery executing " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Property pnEventId() As Integer
        Get
            Dim o As Object = ViewState("MP_EventId")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("MP_EventId") = Value
        End Set
    End Property
    
    Property pnDisplayMode() As Integer
        Get
            Dim o As Object = ViewState("MP_DisplayMode")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("MP_DisplayMode") = Value
        End Set
    End Property

    Protected Sub lnkbtnManageCreditLimits_Click(sender As Object, e As System.EventArgs)
        If trHardieManageCredit02.Visible = False Then
            trHardieManageCredit02.Visible = True
            trHardieManageCredit03.Visible = True
            trHardieManageCredit04.Visible = True
            trHardieManageCredit05.Visible = True
            Call InitCreditLimit()
        End If
    End Sub
    
    Protected Sub InitCreditLimit()
        Dim sSQL As String = "SELECT [key] 'UserKey', FirstName, LastName, UserID FROM UserProfile WHERE Status = 'Active' AND DeletedFlag = 0 AND CustomerKey = " & Session("CustomerKey") & " AND NOT [key] IN (22909, 23138, 23123) ORDER BY LastName"
        Dim dtUsers As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlHardieUsers.Items.Clear()
        ddlHardieUsers.Items.Add(New ListItem("- please select -", 0))
        For Each drUser As DataRow In dtUsers.Rows
            ddlHardieUsers.Items.Add(New ListItem(drUser("FirstName") & " " & drUser("LastName") & " (" & drUser("UserID") & ")", drUser("UserKey")))
        Next
        Dim dtCredit As DataTable
        sSQL = "SELECT MonthOrderLimit FROM ClientData_JamesHardie_MonthOrderLimit WHERE UserKey = 0"
        dtCredit = ExecuteQueryToDataTable(sSQL)
        If dtCredit.Rows.Count = 1 Then
            tbDefaultCreditLimit.Text = Format(dtCredit.Rows(0).Item(0), "#.##")
        Else
            tbDefaultCreditLimit.Text = "UNDEFINED !!"
        End If
        Call BindUserCreditLimits()
    End Sub

    Protected Sub BindUserCreditLimits()
        'Dim sSQL As String = "SELECT up.[key] 'UserKey', FirstName + ' ' + LastName + ' (' + UserID + ')' 'User', MonthOrderValue , ISNULL(CAST (MonthOrderLimit AS varchar(10)), 'default') 'CreditLimit', OverrideCheck FROM ClientData_JamesHardie_MonthOrderLimit mol RIGHT OUTER JOIN ClientData_JamesHardie_MonthOrderValue mov ON mol.UserKey = mov.UserKey INNER JOIN UserProfile up ON mov.UserKey = up.[key] WHERE (NOT up.[key] IN (23123, 23138)) AND up.CustomerKey = " & Session("CustomerKey") & " ORDER BY up.LastName"
        Dim sSQL As String = "SELECT up.[key] 'UserKey', FirstName + ' ' + LastName + ' (' + UserID + ')' 'User', MonthOrderValue , ISNULL(CAST (MonthOrderLimit AS varchar(10)), 'default') 'CreditLimit', OverrideCheck FROM ClientData_JamesHardie_MonthOrderLimit mol RIGHT OUTER JOIN ClientData_JamesHardie_MonthOrderValue mov ON mol.UserKey = mov.UserKey RIGHT OUTER JOIN UserProfile up ON mov.UserKey = up.[key] WHERE (NOT up.[key] IN (22909, 23123, 23138)) AND up.CustomerKey = " & Session("CustomerKey") & " ORDER BY up.LastName"
        Dim dtCredit As DataTable = ExecuteQueryToDataTable(sSQL)
        gvUserCreditLimits.DataSource = dtCredit
        gvUserCreditLimits.DataBind()
    End Sub
    
    Protected Sub btnSaveDefaultCreditLimit_Click(sender As Object, e As System.EventArgs)
        tbDefaultCreditLimit.Text = tbDefaultCreditLimit.Text.Trim
        If IsNumeric(tbDefaultCreditLimit.Text) AndAlso CInt(tbDefaultCreditLimit.Text) > 0 Then
            If CInt(tbDefaultCreditLimit.Text) < 5000 Then
                Call ExecuteQueryToDataTable("UPDATE ClientData_JamesHardie_MonthOrderLimit SET MonthOrderLimit = " & CInt(tbDefaultCreditLimit.Text) & " WHERE UserKey = 0")
            Else
                WebMsgBox.Show("Default credit limit must be less than £5000.")
            End If
        Else
            WebMsgBox.Show("Default credit limit must be a positive number.")
        End If
    End Sub

    Protected Sub btnSaveIndividualCreditLimit_Click(sender As Object, e As System.EventArgs)
        tbIndividualCreditLimit.Text = tbIndividualCreditLimit.Text.Trim
        If tbIndividualCreditLimit.Text <> String.Empty And cbNoCreditLimitCheck.Checked Then
            WebMsgBox.Show("Credit limit must be left blank if No check is selected.")
            Exit Sub
        End If
        If ddlHardieUsers.SelectedValue > 0 Then
            If cbNoCreditLimitCheck.Checked Then
                Dim sSQL As String = "IF EXISTS (SELECT * FROM ClientData_JamesHardie_MonthOrderValue WHERE UserKey = " & ddlHardieUsers.SelectedValue & ") UPDATE ClientData_JamesHardie_MonthOrderValue SET OverrideCheck = 1 WHERE UserKey = " & ddlHardieUsers.SelectedValue & " ELSE INSERT INTO ClientData_JamesHardie_MonthOrderValue (UserKey, MonthOrderValue, LastUpdatedOn, OverrideCheck) VALUES (" & ddlHardieUsers.SelectedValue & ", 0, GETDATE(), 1)"
                Call ExecuteQueryToDataTable(sSQL)
                Call BindUserCreditLimits()
            Else
                If IsNumeric(tbIndividualCreditLimit.Text) AndAlso CInt(tbIndividualCreditLimit.Text) > 0 Then
                    If CInt(tbIndividualCreditLimit.Text) < 5000 Then
                        Call ExecuteQueryToDataTable("DELETE FROM ClientData_JamesHardie_MonthOrderLimit WHERE UserKey = " & ddlHardieUsers.SelectedValue & " UPDATE ClientData_JamesHardie_MonthOrderValue SET OverrideCheck = 0 WHERE UserKey = " & ddlHardieUsers.SelectedValue & " INSERT INTO ClientData_JamesHardie_MonthOrderLimit (UserKey, MonthOrderLimit) VALUES (" & ddlHardieUsers.SelectedValue & ", " & CInt(tbIndividualCreditLimit.Text) & ") IF NOT EXISTS (SELECT * FROM ClientData_JamesHardie_MonthOrderValue WHERE UserKey = " & ddlHardieUsers.SelectedValue & ") INSERT INTO ClientData_JamesHardie_MonthOrderValue (UserKey, MonthOrderValue, LastUpdatedOn, OverrideCheck) VALUES (" & ddlHardieUsers.SelectedValue & ", 0, GETDATE(), 0)")
                    Else
                        WebMsgBox.Show("Credit limit must be less than £5000.")
                    End If
                Else
                    WebMsgBox.Show("Credit limit must be a positive number.")
                End If
            End If
        Else
            WebMsgBox.Show("Please choose a user.")
        End If
        ddlHardieUsers.SelectedIndex = 0
        tbIndividualCreditLimit.Text = String.Empty
        cbNoCreditLimitCheck.Checked = False
        Call BindUserCreditLimits()
    End Sub
    
    Protected Sub gvUserCreditLimits_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvrea As GridViewRowEventArgs = e
        Dim row As GridViewRow = gvrea.Row
        If row.RowType = DataControlRowType.DataRow Then
            Dim sCreditLimitText As String = row.Cells(3).Text
            If sCreditLimitText.Contains("default") Then
                'row.Cells(0).Visible = False
                Dim lnkbtn As LinkButton = row.Cells(0).FindControl("lnkbtnDeleteIndividualCreditEntry")
                lnkbtn.Visible = False
            End If
        End If
    End Sub

    Protected Sub lnkbtnDeleteIndividualCreditEntry_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nUserKey As Int32 = lnkbtn.CommandArgument
        Call ExecuteQueryToDataTable("DELETE FROM ClientData_JamesHardie_MonthOrderLimit WHERE UserKey = " & nUserKey & " UPDATE ClientData_JamesHardie_MonthOrderValue SET OverrideCheck = 0 WHERE UserKey = " & nUserKey)
        Call BindUserCreditLimits()
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>My Profile</title>
    <link href="sprint.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .style1
        {
            color: #800000;
        }
    </style>
    </head>
    <body>
        <form id="Form1" runat="Server">
        <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_myprofile">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                      <strong style="color: navy; font-size:x-small; font-family:Verdana">My Profile</strong>
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
            <asp:panel ID="pnlProfile" runat="server" Width="100%">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%" align="right">
                        <strong>Contact information:</strong></td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                        <strong>
                        <asp:Label ID="lblLegendSecurityQuestions" runat="server" 
                            Text="Security questions:" Visible="False" />
                        </strong></td>
                    <td style="width: 29%" align="right">
                        <asp:LinkButton ID="lnkbtnChangeLoginPassword" runat="server" OnClick="lnkbtnChangeLoginPassword_Click">change login password</asp:LinkButton></td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        Telephone:</td>
                    <td>
                        <asp:TextBox ID="tbTelephone" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="200px" MaxLength="20"></asp:TextBox></td>
                    <td align="right">
                        <asp:Label ID="lblLegendBirthday" runat="server" 
                            Text="Birthday of your spouse / partner / significant other (dd/mm/yy)?" 
                            Visible="False" />
                        </td>
                    <td>
                        <asp:TextBox ID="tbMemorableAnswer1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px" Visible="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        Email:</td>
                    <td>
                        <asp:TextBox ID="tbEmail" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px" MaxLength="100"></asp:TextBox></td>
                    <td align="right">
                        <asp:Label ID="lblLegendStreetNumberOfTheHouseYouFirstLivedIn" runat="server" 
                            Text="Street number of the house you first lived in?" Visible="False" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbMemorableAnswer2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px" Visible="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        Cost Centre:</td>
                    <td>
                        <asp:TextBox ID="tbCostCentre" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Width="200px" MaxLength="20"></asp:TextBox></td>
                    <td align="right">
                        <asp:Label ID="lblLegendGrandmothersMaidenName" runat="server" Text="Grandmother's maiden name?" Visible="False" />
                    </td>
                    <td>
                        <asp:TextBox ID="tbMemorableAnswer3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="50" Width="200px" Visible="False" /></td>
                    <td>
                    </td>
                </tr>
                <tr id="trDefaultDestination" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        Default Destination:</td>
                    <td colspan="3">
                        <asp:Label ID="lblDefaultDestination" runat="server" ForeColor="Silver"></asp:Label>&nbsp;
                        <asp:LinkButton ID="lnkbtnCancelDefaultDestination" runat="server" OnClick="lnkbtnCancelDefaultDestination_Click">cancel default</asp:LinkButton></td>
                    <td>
                    </td>
                </tr>
                <tr id="trAuthorisationExempt" runat="server">
                    <td>
                    </td>
                    <td align="right">
                        Authorisation status:</td>
                    <td colspan="3">
                        <asp:Label ID="Label21" runat="server" Font-Bold="True" ForeColor="#00C000" Text="You are exempt from authorisation"></asp:Label></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td colspan="3">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr id="trHardieMonthlyCreditRemaining" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:Label ID="lblLegendHardieCreditRemaining" runat="server" Font-Bold="True" 
                            Text="Order credit remaining for this month:" Font-Size="X-Small" />
                    </td>
                    <td colspan="3">
                        <asp:Label ID="lblHardieCreditRemaining" runat="server" Font-Bold="True" 
                            Font-Size="X-Small" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit01" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:LinkButton ID="lnkbtnManageCreditLimits" runat="server" 
                            onclick="lnkbtnManageCreditLimits_Click">manage credit limits</asp:LinkButton>
                    </td>
                    <td colspan="3">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit02" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        Default credit limit (£):</td>
                    <td colspan="3">
                        <asp:TextBox ID="tbDefaultCreditLimit" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="5" Width="100px"/>
                        &nbsp;<asp:Button ID="btnSaveDefaultCreditLimit" runat="server" Text="save" 
                            onclick="btnSaveDefaultCreditLimit_Click" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit03" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        Add /change individual credit limit:</td>
                    <td colspan="3">
                        <asp:DropDownList ID="ddlHardieUsers" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small"/>
                        &nbsp; Credit limit (£):
                        <asp:TextBox ID="tbIndividualCreditLimit" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="5" Width="100px" />
                        &nbsp;<asp:CheckBox ID="cbNoCreditLimitCheck" runat="server" Text="No check" />
                        &nbsp;
                        <asp:Button ID="btnSaveIndividualCreditLimit" runat="server" Text="save" 
                            onclick="btnSaveIndividualCreditLimit_Click" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit04" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        Individual credit limits:</td>
                    <td colspan="3">
                        <asp:GridView ID="gvUserCreditLimits" OnRowDataBound="gvUserCreditLimits_RowDataBound" runat="server" CellPadding="1" 
                            AutoGenerateColumns="False">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnDeleteIndividualCreditEntry" 
                                            CommandArgument='<%# Container.DataItem("UserKey")%>' runat="server" 
                                            onclick="lnkbtnDeleteIndividualCreditEntry_Click">set to default</asp:LinkButton>
                                        <%--<asp:HiddenField ID="hidUserKey" Value='<%# Container.DataItem("UserKey")%>' runat="server" />--%>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center" Width="100px" />
                                </asp:TemplateField>
                                <asp:BoundField DataField="User" HeaderText="User" ReadOnly="True" 
                                    SortExpression="User" />
                                <asp:BoundField DataField="MonthOrderValue" HeaderText="&nbsp;Month Order Value&nbsp;" 
                                    ReadOnly="True" SortExpression="MonthOrderValue" DataFormatString="{0:C}">
                                <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="CreditLimit" HeaderText="&nbsp;Credit Limit&nbsp;" 
                                    ReadOnly="True" SortExpression="CreditLimit" DataFormatString="{0:C}">
                                <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                                <asp:BoundField DataField="OverrideCheck" HeaderText="&nbsp;No Check&nbsp;" ReadOnly="True" 
                                    SortExpression="OverrideCheck" >
                                <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundField>
                            </Columns>
                        </asp:GridView>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit05" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td colspan="3" class="style1">
                        NOTE: The monthly order total may not be reset until the first order of a new 
                        month is placed.</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr ID="trHardieManageCredit06" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td colspan="3">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
            <table id="tblAlerts" runat="server" style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <strong>EMAIL ALERTS &amp; CONFIRMATIONS</strong>
                    </td>
                    <td>
                        <asp:CheckBox ID="cbReceiveOrderConfirmationAlertsUser" runat="server" 
                            Text="Receive confirmation of my orders" /></td>
                    <td>
                        </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trReceiveOrderConfirmationAlerts" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbReceiveOrderConfirmationAlertsSuperUser" runat="server" 
                            Text="Receive confirmation of ALL orders" />
                        <asp:CheckBox ID="cbReceiveOrderConfirmationAlertsProductOwner" runat="server" 
                            Text="Receive confirmation of ALL orders  for products I manage" /></td>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trReceiveGoodsInAlerts" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbReceiveGoodsInAlertsSuperUser" runat="server" 
                            Text="Receive ALL Goods In alerts" />
                        <asp:CheckBox ID="cbReceiveGoodsInAlertsProductOwner" runat="server" 
                            Text="Receive Goods In alerts For products I manage" /></td>
                    <td>
                        </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trReceiveLowStockAlerts" runat="server" visible="false">
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbReceiveLowStockAlertsSuperUser" runat="server" 
                            Text="Receive ALL Low Stock alerts" />
                        <asp:CheckBox ID="cbReceiveLowStockAlertsProductOwner" runat="server" 
                            Text="Receive Low Stock alerts for products I manage" /></td>
                    <td>
                        </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trLovellsUser" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td colspan="3">
                        <asp:CheckBox ID="cbReceivePublicationOrderAlerts" runat="server" 
                            Text="Receive alerts when my publications are ordered" />
                        <br />
                        <asp:CheckBox ID="cbReceiveProductInactivityAlerts" runat="server" 
                            Text="Receive product inactivity alerts for my publications" />
                        <br />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
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
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:Button ID="btnSaveChanges" runat="server" OnClick="btnSaveChanges_Click" Text="save profile changes" /></td>
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
                    <td align="right">
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
            </table>
            </asp:panel>
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                      <asp:LinkButton ID="lnkbtnShowHideProfile" runat="server" OnClick="lnkbtnShowHideProfile_Click">hide profile</asp:LinkButton>
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
            </table>
            <asp:Panel ID="pnlCalendarManagement" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 29%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                      <asp:Label ID="lblLegendCalendarManagedItems" runat="server" Font-Bold="True" Text="Calendar Managed Products"></asp:Label>
                    </td>
                    <td align="right" colspan="2">
                        <asp:LinkButton ID="lnkbtnCMHelp" runat="server" OnClientClick='window.open("help_cmproductsuser.pdf", "CMHelp","top=10,left=10,width=700,height=450,status=no,toolbar=yes,address=no,menubar=yes,resizable=yes,scrollbars=yes");'>calendar managed products help</asp:LinkButton></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:Button ID="btnCMShowMyEvents" runat="server" Text="show my events" OnClick="btnCMShowEvents_Click" CommandArgument="my" />
                        <asp:Button ID="btnCMShowAllEvents" runat="server" Text="show all events" OnClick="btnCMShowEvents_Click" CommandArgument="all" />
                        &nbsp;
                        <asp:RadioButton ID="rbShowByEvent" runat="server" Checked="True" Text="by event" GroupName="CM1" OnCheckedChanged="rbShowByEvent_CheckedChanged" AutoPostBack="True" />
                        <asp:RadioButton ID="rbShowByItem" runat="server" GroupName="CM1" Text="by product" OnCheckedChanged="rbShowByItem_CheckedChanged" AutoPostBack="True" />
                        &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;
                        <asp:Label ID="lblCMLegendSelectItem" runat="server" Text="product:" 
                            Visible="False"></asp:Label>
                        <asp:DropDownList ID="ddlCMItems" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlCMItems_SelectedIndexChanged" Visible="False" AutoPostBack="True">
                        </asp:DropDownList></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <asp:GridView ID="gvCalendarManagedItems" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" OnRowDataBound="gvCalendarManagedItems_RowDataBound" OnPageIndexChanging="gvCalendarManagedItems_PageIndexChanging" AllowPaging="True">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnCMShowDetails" CommandArgument='<%# Container.DataItem("EventId")%>' CommandName='<%# gvCMSetButtonVisibility(Container.DataItem) %>' runat="server" OnClick="lnkbtnCMShowDetails_Click">show</asp:LinkButton>
                                        <asp:LinkButton ID="lnkbtnCMDeleteEvent" CommandArgument='<%# Container.DataItem("EventId")%>' Visible='<%# gvCMSetButtonVisibility(Container.DataItem) %>' OnClick="lnkbtnCMDeleteEvent_Click" OnClientClick='return confirm("Are you sure you want to delete this event?");' runat="server">delete</asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <EmptyDataTemplate>
                                no items found
                            </EmptyDataTemplate>
                            <PagerStyle HorizontalAlign="Center" />
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
            </asp:Panel>
        <asp:Panel ID="pnlEvent" runat="server" Width="100%" Visible="False">
            <table style="width: 100%">
                <tr>
                    <td style="width: 1%">
                    </td>
                    <td style="width: 16%" align="right">
                    </td>
                    <td style="width: 33%">
                    </td>
                    <td style="width: 16%">
                    </td>
                    <td style="width: 33%">
                    </td>
                    <td style="width: 1%">
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        &nbsp;<asp:Label ID="lblLegendEvent" runat="server" Text="Event Details:" Font-Bold="True"/></td>
                    <td>
                    </td>
                    <td style="width: 450px">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label1" runat="server" Visible="false" Font-Names="Verdana" Font-Size="XX-Small" Text="Customer:"/></td>
                    <td>
                        <asp:Label ID="lblCustomer" runat="server" Visible="false" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"/></td>
                    <td colspan="2">
                        <asp:Label ID="Label16" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Booked by"/>
                        <asp:Label ID="lblBookedBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/><asp:Label ID="Label17" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="&nbsp;on"/>
                        <asp:Label ID="lblBookedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Event name:"/></td>
                    <td>
                        <asp:Label ID="lblEventName" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"/></td>
                    <td colspan="2">
                        <asp:Label ID="Label20" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Delivery date:"/>
                        <asp:Label ID="lblDeliveryDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"/>
                        &nbsp;
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection date:"/>
                        <asp:Label ID="lblCollectionDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvContactName" ControlToValidate="tbContactName" runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact name:" ForeColor="Red"/>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="tbContactName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvContactPhone" ControlToValidate="tbContactPhone" runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact phone:" ForeColor="Red"/></td>
                    <td>
                        <asp:TextBox ID="tbContactPhone" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvContactMobile" ControlToValidate="tbContactMobile" runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Contact mobile:" ForeColor="Red"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbContactMobile" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactName2" runat="server" ControlToValidate="tbCMContactName2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" Visible="False"/>
                        <asp:Label ID="lblLegendCMContactName2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact name 2:" Visible="False"/>
                    </td>
                    <td>
                        <asp:LinkButton ID="lnkbtnCMAddSecondContact" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnCMAddSecondContact_Click">add 2nd contact</asp:LinkButton>                    
                        <asp:TextBox ID="tbCMContactName2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" Visible="False" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactMobile2" runat="server" ControlToValidate="tbCMContactMobile2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="lblLegendCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Contact mobile 2:"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCMContactMobile2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td>
                        &nbsp;
                    </td>
                </tr>
                <tr ID="trCMContactPhone2" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMContactPhone2" runat="server" ControlToValidate="tbCMContactPhone2" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label15axa0" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Contact phone 2:"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCMContactPhone2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" MaxLength="50" Width="300px" 
                            BackColor="LightGoldenrodYellow" />
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        <asp:LinkButton ID="lnkbtnCMRemoveSecondContact" runat="server" 
                            OnClientClick='return confirm("Are you sure you want to remove the 2nd contact?");' 
                            onclick="lnkbtnCMRemoveSecondContact_Click">remove 2nd contact</asp:LinkButton>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td >
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvEventAddress1" runat="server" 
                            ControlToValidate="tbEventAddress1" ErrorMessage="#" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="Event address 1:" ForeColor="Red"/></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbEventAddress1" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="400px" MaxLength="50" 
                            BackColor="LightGoldenrodYellow"/></td>
                    <td>
                        </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="Event address 2:"/></td>
                    <td>
                        <asp:TextBox ID="tbEventAddress2" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="400px" MaxLength="50" 
                            BackColor="LightGoldenrodYellow"/></td>
                    <td align="right">
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="Event address 3:"/></td>
                    <td>
                        <asp:TextBox ID="tbEventAddress3" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="400px" MaxLength="50" 
                            BackColor="LightGoldenrodYellow"/></td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvTown" runat="server" 
                            ControlToValidate="tbTown" ErrorMessage="#" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="Town:" ForeColor="Red"/></td>
                    <td>
                        <asp:TextBox ID="tbTown" runat="server" BackColor="LightGoldenrodYellow" 
                            Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="300px" />
                    </td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvPostCode" 
                            ControlToValidate="tbPostcode" runat="server" ErrorMessage="#" 
                            ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="Post code:" ForeColor="Red"/></td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPostcode" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="150px" MaxLength="50" 
                            BackColor="LightGoldenrodYellow"/>&nbsp;<asp:LinkButton 
                            ID="lnkbtnCMAddressOutsideUK" runat="server" 
                            onclick="lnkbtnCMAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trCMCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCountry" runat="server" 
                            ControlToValidate="ddlCMCountry" ErrorMessage="#" Font-Names="Verdana" 
                            Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa0" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCountry" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px">
                        </asp:DropDownList>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td/>
                    <td align="right">
                        <asp:Label ID="Label12" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Delivery time:" ForeColor="Red"/>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlDeliveryTime" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" BackColor="LightGoldenrodYellow">
                            <asp:ListItem>9.00am</asp:ListItem>
                            <asp:ListItem>10.30am</asp:ListItem>
                            <asp:ListItem>12.00 noon</asp:ListItem>
                            <asp:ListItem>Other times pls specify in Special Instructions</asp:ListItem>
                        </asp:DropDownList></td>
                    <td align="right"><asp:RequiredFieldValidator ID="rfvPreciseDeliveryPoint" ControlToValidate="tbPreciseDeliveryPoint" runat="server" ErrorMessage="#" ValidationGroup="CalendarManaged"/>
                        <asp:Label ID="Label13" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Exact delivery point:" ForeColor="Red"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPreciseDeliveryPoint" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" MaxLength="100" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td/>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Collection time:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCollectionTime" runat="server" 
                            BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small">
                            <asp:ListItem>9.00am - 10.00am</asp:ListItem>
                            <asp:ListItem>10.00am - 11.00am</asp:ListItem>
                            <asp:ListItem>11.00am - 12.00 noon</asp:ListItem>
                            <asp:ListItem>12.00 noon - 1.00pm</asp:ListItem>
                            <asp:ListItem>1.00pm - 2.00pm</asp:ListItem>
                            <asp:ListItem>2.00pm - 3.00pm</asp:ListItem>
                            <asp:ListItem>3.00pm - 4.00pm</asp:ListItem>
                            <asp:ListItem>4.00pm - 5.00pm</asp:ListItem>
                            <asp:ListItem>5.00pm - 6.00pm</asp:ListItem>
                            <asp:ListItem>Other - contact Sprint</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvPreciseCollectionPoint" runat="server" ControlToValidate="tbPreciseCollectionPoint" ErrorMessage="#" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Exact collection point:" />
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbPreciseCollectionPoint" runat="server" BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="100" Width="100%" />
                    </td>
                    <td/>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td>
                        <asp:CheckBox ID="cbDifferentCollectionAddress" runat="server" 
                            AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" 
                            OnCheckedChanged="cbDifferentCollectionAddress_CheckedChanged" 
                            Text="collect from a different address" />
                    </td>
                    <td align="right">
                        </td>
                    <td style="width: 450px">
                    </td>
                    <td/>
                </tr>
                <tr id="trCollectionAddress1" runat="server" visible="false">
                    <td style="height: 18px">
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCollectionAddress1" runat="server" 
                            ControlToValidate="tbCollectionAddress1" ErrorMessage="#" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label22" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection address 1:" ForeColor="Red"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCollectionAddress1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label23" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection address 2:"/>
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCollectionAddress2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="400px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td/>
                </tr>
                <tr ID="trCollectionAddress2" runat="server" visible="false">
                    <td/>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCollectionTown" runat="server" 
                            ControlToValidate="tbCollectionTown" ErrorMessage="#" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label24" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Collection town:" ForeColor="Red"/>
                    </td>
                    <td>
                        <asp:TextBox ID="tbCollectionTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="300px" MaxLength="50" BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvPostCode0" runat="server" 
                            ControlToValidate="tbCollectionPostcode" ErrorMessage="#" 
                            ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label39" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Collection post code:" />
                    </td>
                    <td style="width: 450px">
                        <asp:TextBox ID="tbCollectionPostcode" runat="server" BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="150px" />
                        &nbsp;<asp:LinkButton ID="lnkbtnCMCollectionAddressOutsideUK" runat="server" 
                            onclick="lnkbtnCMCollectionAddressOutsideUK_Click">addr outside UK</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr ID="trCMCollectionCountry" runat="server" visible="false">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:RequiredFieldValidator ID="rfvCMCollectionCountry" runat="server" ControlToValidate="ddlCMCollectionCountry" ErrorMessage="#" Font-Names="Verdana" Font-Size="XX-Small" InitialValue="0" ValidationGroup="CalendarManaged" />
                        <asp:Label ID="Label38axa1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red" Text="Country:" />
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlCMCollectionCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy" TabIndex="8" Width="250px"/>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td style="width: 450px">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label40" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="Customer reference:"/>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="tbCustomerReference" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Width="400px" MaxLength="50" 
                            BackColor="LightGoldenrodYellow"/>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label19" runat="server" Text="Special instructions:" 
                            Font-Names="Verdana" Font-Size="XX-Small"/></td>
                    <td colspan="3">
                        <asp:TextBox ID="tbSpecialInstructions" runat="server" 
                            BackColor="LightGoldenrodYellow" Font-Names="Verdana" Font-Size="XX-Small" 
                            MaxLength="180" Width="99%" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="lblLegendProduct" runat="server" Text="Product:" />
                    </td>
                    <td colspan="3">
                        <asp:GridView ID="gvItems" runat="server" CellPadding="2" Width="100%">
                        </asp:GridView>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3">
                        <asp:Button ID="btnSaveEventChanges" runat="server" OnClick="btnSaveEventChanges_Click" Text="save event changes" />
                        <asp:Label ID="lblOnlineChangesMessage" runat="server" Text="No changes can be accepted online as there are 5 days or fewer remaining until delivery. Contact Customer Services to request changes." Visible="False" ForeColor="Red"></asp:Label>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="4">
                        <hr />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr ID="trNotes" runat="server">
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3">
                        <asp:Label ID="Label25" runat="server" Font-Bold="True" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="Notes:" />
                        <asp:GridView ID="gvNotes" runat="server" AllowPaging="True" CellPadding="2" 
                            OnPageIndexChanging="gvNotes_PageIndexChanging" 
                            OnRowDataBound="gvNotes_RowDataBound" PageSize="6" Width="100%">
                            <EmptyDataTemplate>
                                no notes
                            </EmptyDataTemplate>
                            <PagerStyle Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                        </asp:GridView>
                        &nbsp; </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                    </td>
                    <td colspan="3">
                        <asp:LinkButton ID="lnkbtnShowHideNotes" runat="server" 
                            OnClick="lnkbtnShowHideNotes_Click">hide notes</asp:LinkButton>
                        <asp:LinkButton ID="lnkbtnRefreshNotes" runat="server" 
                            OnClick="lnkbtnRefreshNotes_Click">refresh notes</asp:LinkButton>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
        </asp:Panel>
            <asp:Panel ID="pnlAuthorisations" runat="server" Width="100%" Visible="false">
            <table style="width: 100%; font-family:Verdana; font-size:xx-small">
                <tr>
                    <td style="width: 10%; height: 14px;">
                    </td>
                    <td style="width: 20%; height: 14px;">
                            </td>
                    <td style="width: 70%; height: 14px;">
                        </td>
                </tr>
                <tr>
                    <td colspan="3">
                    <hr />
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="lblAuthorisationMessage" runat="server" Font-Bold="True" Font-Names="Verdana"
                            Font-Size="XX-Small">Orders Awaiting Authorisation</asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:GridView ID="gvAuthoriseOrder" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            GridLines="None" Width="100%" AllowPaging="True" AutoGenerateColumns="False">
                            <PagerStyle HorizontalAlign="Center" />
                        <Columns>
                        <asp:TemplateField HeaderText="Order Created" >
                            <ItemTemplate>
                              <asp:Label ID="lblOrderCreatedDateTime" runat="server" Text='<%# Format(Container.DataItem("OrderCreatedDateTime"),"dd-MMM-yy hh:mm:ss")%>' ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Consignee" >
                            <ItemTemplate>
                              <asp:Label ID="lblCneeName" runat="server" Text='<%# Container.DataItem("CneeName")%>' ></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="" >
                            <ItemTemplate>
                                <asp:Button ID="btnViewOrder" runat="server" CommandArgument='<%# Container.DataItem("id")%>' Text="view" OnClick="btnViewOrder_Click" />&nbsp;
                                <asp:Button ID="btnDeleteOrder" runat="server" CommandArgument='<%# Container.DataItem("id")%>' Text="delete" OnClientClick='return confirm("Are you sure you want to delete this order?");' OnClick="btnDeleteOrder_Click" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        </Columns>
                        <EmptyDataTemplate>
                            no orders are awaiting authorisation
                        </EmptyDataTemplate>
                        <RowStyle BackColor="WhiteSmoke" />
                        <AlternatingRowStyle BackColor="White" />
                        </asp:GridView>
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
            </asp:Panel>

        <asp:Panel ID="pnlOrderDetail" runat="server" Visible="False" Width="100%">
            <hr />
            <asp:Label ID="Label4" runat="server" Font-Size="xX-Small" Font-Bold="true" Font-Names="Verdana" Text="The following order is awaiting authorisation"></asp:Label>
            <br />
            <br />
            <table width="95%">
                <tr>
                    <td style="width: 20%">
                        <asp:Label ID="Label27" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Ordered by:"></asp:Label>
                    </td>
                    <td style="width: 80%">
                        <asp:Label ID="lblAuthOrderOrderedBy" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="height: 18px">
                        <asp:Label ID="Label26" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Order created on:"></asp:Label>
                    </td>
                    <td style="height: 18px">
                        <asp:Label ID="lblAuthOrderPlacedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
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
                        <asp:Label ID="Label30" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" Text="CONSIGNEE DETAILS"></asp:Label>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label28" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Consignee:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderConsignee" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label29" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Attn Of:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderAttnOf" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="height: 20px">
                        <asp:Label ID="Label31" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Addr 1:"></asp:Label>
                    </td>
                    <td style="height: 20px">
                        <asp:Label ID="lblAuthOrderAddr1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label32" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Addr 2:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderAddr2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label34" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Addr 3:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderAddr3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label33" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Town/City:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderTown" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label35" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="County/State"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderState" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label36" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Postcode:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderPostcode" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label37" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Country:"></asp:Label>
                    </td>
                    <td>
                        <asp:Label ID="lblAuthOrderCountry" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
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
                        <asp:Label ID="Label38" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" Text="CONSIGNMENT DETAILS" /><asp:Label ID="Label41" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=" (items in " /><asp:Label ID="Label42" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="red" Text="red" /><asp:Label ID="Label43" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=" require authorisation)" />
                    </td>
                </tr>
            </table>
            <br />
            <asp:GridView ID="gvAuthOrderDetails" runat="server"  Font-Names="Verdana" Font-Size="XX-Small" CellPadding="3"
                 AutoGenerateColumns="False" Width="100%" GridLines="None" >
                <Columns>
                    <asp:TemplateField HeaderText="Product Code" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductCodeView" runat="server" Text='<%# Container.DataItem("ProductCode")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Product Date" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDateView" runat="server" Text='<%# Container.DataItem("ProductDate")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Description" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDescriptionView" runat="server" Text='<%# Container.DataItem("ProductDescription")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductItemsOutView" runat="server" Text='<%# Container.DataItem("ItemsOut")%>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <EmptyDataTemplate>
                    no items found in order
                </EmptyDataTemplate>
                <RowStyle BackColor="WhiteSmoke" />
                <AlternatingRowStyle BackColor="White" />
            </asp:GridView>
            <br />
            <asp:Button ID="btnGoBack" runat="server" Text="go back" OnClick="btnGoBack_Click" />&nbsp;
            <asp:Button ID="btnDeleteOrderFromDetailScreen" runat="server" Text="delete" OnClientClick='return confirm("Are you sure you want to delete this order?");' OnClick="btnDeleteOrderFromDetailScreen_Click" />
            <asp:HiddenField ID="hidHoldingQueueKey" runat="server" />
        </asp:Panel>

            <br />
            <table id="tblProductCodeReservations" runat="server" visible="false" style="width: 100%; font-family:Verdana; font-size:xx-small">
                <tr>
                    <td style="width: 10%">
                    </td>
                    <td colspan="2">
                        <asp:Label ID="lblReservationMessage" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;
                    </td>
                    <td colspan="2">
                        <asp:GridView ID="gvProductCodeReservations" runat="server" Font-Names="Verdana"
                            Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" CellPadding="3">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkbtnCancelReservation" runat="server" CommandArgument='<%# Eval("id") %>' OnClick="lnkbtnCancelReservation_Click">cancel reservation</asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Product Code" SortExpression="ProductCode">
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("ProductCode") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Date Reserved" SortExpression="DateReserved">
                                    <ItemTemplate>
                                        <asp:Label ID="Label2" runat="server" Text='<%# Format(Eval("DateReserved"), "d-MMM-yy") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Reason" SortExpression="Notes">
                                    <ItemTemplate>
                                        <asp:Label ID="Label3" runat="server" Text='<%# Eval("Notes") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <AlternatingRowStyle BackColor="#FBFBFB" />
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;
                    </td>
                    <td colspan="2">
                        </td>
                </tr>
                <tr>
                    <td>&nbsp;
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>&nbsp;
                    </td>
                    <td>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
    </form>
</body>
</html>

