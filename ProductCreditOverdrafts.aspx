<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

    ' TO DO
   
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call HideAllPanelsAndRows()
            Call BindOverdraftRequests()
        End If
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Product Credit Overdrafts"
    End Sub
   
    Protected Sub HideAllPanelsAndRows()
        pnlOverdraftRequests.Visible = False
        'gvOverdraftRequests.Visible = False
        pnlOverdraftRequest.Visible = False
    End Sub
    
    Protected Sub BindOverdraftRequests()
        Dim sIncludeTypes As String = "'queued'"
        If cbIncludeAuthorisedRequests.Checked Then
            sIncludeTypes &= ",'authorised'"
        End If
        If cbIncludeDeclinedRequests.Checked Then
            sIncludeTypes &= ",'declined'"
        End If
        sIncludeTypes = "(" & sIncludeTypes & ")"
        Dim sSQL As String = "SELECT [id], CAST(REPLACE(CONVERT(VARCHAR(11),  OrderCreatedDateTime, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), OrderCreatedDateTime, 108)),1,5) 'OrderCreatedDateTime', FirstName + ' ' + LastName + ' (' + UserID + ')' AgentName, ohq.UserProfileKey, OrderStatus FROM ProductCreditsOrderHoldingQueue ohq INNER JOIN UserProfile up ON ohq.UserProfileKey = up.[key] WHERE OrderStatus IN " & sIncludeTypes & " AND OrderCreatedDateTime >= GETDATE() - " & ddlRequestDays.SelectedValue & " ORDER BY [id]"
        Dim dtOverdraftRequests As DataTable = ExecuteQueryToDataTable(sSQL)
        gvOverdraftRequests.DataSource = dtOverdraftRequests
        gvOverdraftRequests.DataBind()
        'gvOverdraftRequests.Visible = True
        pnlOverdraftRequests.Visible = True
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
    
    Protected Sub btnView_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlOverdraftRequest.Visible = True
        Dim b As Button = sender
        Dim nProductCreditsOrderHoldingQueueKey As Int32 = b.CommandArgument
        Call BindRequest(nProductCreditsOrderHoldingQueueKey)
        btnAuthoriseRequest.CommandArgument = nProductCreditsOrderHoldingQueueKey
        btnDeclineRequest.CommandArgument = nProductCreditsOrderHoldingQueueKey
    End Sub
    
    Protected Sub BindRequest(ByVal nProductCreditsOrderHoldingQueueKey As Int32)
        Dim sSQL As String
        sSQL = "SELECT [id], CAST(REPLACE(CONVERT(VARCHAR(11),  OrderCreatedDateTime, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), OrderCreatedDateTime, 108)),1,5) 'OrderCreatedDateTime', FirstName + ' ' + LastName + ' (' + UserID + ')' AgentName, ohq.UserProfileKey, ISNULL(CneeCtcName,'') 'CneeCtcName', ISNULL(CneeName,'') CneeName, ISNULL(CneeAddr1,'') CneeAddr1, ISNULL(CneeAddr2,'') CneeAddr2, ISNULL(CneeTown,'') CneeTown, ISNULL(CneePostCode ,'') CneePostCode, ISNULL(MsgToOrderer,'') 'MsgToOrderer', OrderStatus FROM ProductCreditsOrderHoldingQueue ohq INNER JOIN UserProfile up ON ohq.UserProfileKey = up.[key] WHERE [id] = " & nProductCreditsOrderHoldingQueueKey
        Dim drOverdraftRequest As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        lblRequestRaisedOn.Text = drOverdraftRequest("OrderCreatedDateTime")
        lblAgent.Text = drOverdraftRequest("AgentName") & " " & drOverdraftRequest("CneeAddr1") & " " & drOverdraftRequest("CneeTown")
        tbMessageToOrderer.Text = drOverdraftRequest("MsgToOrderer")
        sSQL = "SELECT lp.LogisticProductKey, ProductCode, ProductDate, ProductDescription, ItemsOut FROM ProductCreditsOrderItemHoldingQueue oihq INNER JOIN LogisticProduct lp ON oihq.LogisticProductKey = lp.LogisticProductKey WHERE ProductCreditsOrderHoldingQueueKey = " & nProductCreditsOrderHoldingQueueKey & " ORDER BY [id]"
        Dim dtOverdraftRequestItems As DataTable = ExecuteQueryToDataTable(sSQL)
        gvOverdraftRequest.DataSource = dtOverdraftRequestItems
        gvOverdraftRequest.DataBind()
        Call HideAllPanelsAndRows()
        pnlOverdraftRequest.Visible = True
        hidContactName.Value = drOverdraftRequest("CneeCtcName")
        hidName.Value = drOverdraftRequest("CneeName")
        hidAddr1.Value = drOverdraftRequest("CneeAddr1")
        hidAddr2.Value = drOverdraftRequest("CneeAddr2")
        hidTown.Value = drOverdraftRequest("CneeTown")
        hidPostcode.Value = drOverdraftRequest("CneePostCode")
        hidSpecialInstructions.Value = ""
        Call BuildDeliveryString()
        If drOverdraftRequest("OrderStatus").ToString.ToLower = "queued" Then
            btnAuthoriseRequest.Visible = True
            btnDeclineRequest.Visible = True
            btnEditDeliveryDetails.Visible = True
        Else
            btnAuthoriseRequest.Visible = False
            btnDeclineRequest.Visible = False
            btnEditDeliveryDetails.Visible = False
        End If
    End Sub
    
    Protected Function sGetSummary(ByVal DataItem As Object) As String
        Dim sSQL As String
        Dim sbSummary As New StringBuilder
        Dim nProductCreditsOrderHoldingQueueKey As Int32 = DataBinder.Eval(DataItem, "id")
        sGetSummary = "Summary details for record " & nProductCreditsOrderHoldingQueueKey.ToString & " go here."

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

        sGetSummary = sbOverdraftDetails.ToString
    End Function
    
    Protected Sub btnAuthoriseRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nProductCreditsOrderHoldingQueueKey As Int32 = b.CommandArgument
        Dim nConsignmentKey As Int32 = 0
        
        If bCheckConsignment() Then
            Dim drProductCreditsOrderHoldingQueue As DataRow = ExecuteQueryToDataTable("SELECT * FROM ProductCreditsOrderHoldingQueue WHERE [id] = " & nProductCreditsOrderHoldingQueueKey.ToString).Rows(0)
            drProductCreditsOrderHoldingQueue("CneeCtcName") = hidContactName.Value
            drProductCreditsOrderHoldingQueue("CneeName") = hidName.Value
            drProductCreditsOrderHoldingQueue("CneeAddr1") = hidAddr1.Value
            drProductCreditsOrderHoldingQueue("CneeAddr2") = hidAddr2.Value
            drProductCreditsOrderHoldingQueue("CneeTown") = hidTown.Value
            drProductCreditsOrderHoldingQueue("CneePostCode") = hidPostcode.Value
            drProductCreditsOrderHoldingQueue("SpecialInstructions") = hidSpecialInstructions.Value
            nConsignmentKey = nSubmitConsignment(drProductCreditsOrderHoldingQueue)
            If nConsignmentKey > 0 Then
                Dim sSQL As String = "UPDATE ProductCreditsOrderHoldingQueue SET OrderStatus = 'AUTHORISED', ConsignmentKey = " & nConsignmentKey.ToString & ", MsgToOrderer = '" & tbMessageToOrderer.Text.Trim.Replace(Environment.NewLine, " ").Replace("'", "''") & "', OrderPlacedDateTime = GETDATE() WHERE [id] = " & nProductCreditsOrderHoldingQueueKey
                Call ExecuteQueryToDataTable(sSQL)
            Else
                WebMsgBox.Show("Unable to create consignment.")
                ' report consignment failure
            End If
            Call HideAllPanelsAndRows()
            Call BindOverdraftRequests()
        End If
    End Sub
    
    Protected Function GetTotalAvailableQty(ByVal sLogisticProductKey As String) As Int32
        GetTotalAvailableQty = ExecuteQueryToDataTable("SELECT Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ")) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation WHERE LogisticProductKey = " & sLogisticProductKey & ") END").Rows(0).Item(0)
    End Function
    
    Protected Function bCheckConsignment() As Boolean
        bCheckConsignment = True
        ' check at least one non-zero qty
        ' check qty available
        Dim bPositiveQtyFound As Boolean = False
        For Each gvr As GridViewRow In gvOverdraftRequest.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim hidLogisticProductKey As HiddenField = gvr.Cells(3).FindControl("hidLogisticProductKey")
                Dim tbQty As TextBox = gvr.Cells(3).FindControl("tbQty")
                tbQty.Text = tbQty.Text.Trim
                If Not IsNumeric(tbQty.Text) Then
                    WebMsgBox.Show("A quantity field is non-numeric.")
                    Return False
                End If
                Dim nQty As Int32 = CInt(tbQty.Text)
                If nQty < 0 Then
                    WebMsgBox.Show("At least one quantity field is negative.")
                    Return False
                End If
                If nQty > 0 Then
                    bPositiveQtyFound = True
                End If
                Dim nTotalAvailableQty = GetTotalAvailableQty(CInt(hidLogisticProductKey.Value))
                If nQty > nTotalAvailableQty Then
                    Dim sProductDetails() As String = GetProductDetailsFromProductKey(CInt(hidLogisticProductKey.Value))
                    WebMsgBox.Show("You have requested " & nQty.ToString & " of product " & sProductDetails(0) & " " & sProductDetails(1) & " - " & sProductDetails(2) & " but the available amount is only " & nTotalAvailableQty.ToString & ".")
                    Return False
                End If
            End If
        Next
        If Not bPositiveQtyFound Then
            WebMsgBox.Show("To place an order you must have one or more items with a quantity greater than zero.")
        End If
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

    Protected Function nSubmitConsignment(ByVal drProductCreditsOrderHoldingQueue As DataRow) As Int32
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        nSubmitConsignment = 0
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        'lblError.Text = ""
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        'param1.Value = ddlBookedBy.SelectedValue
        param1.Value = drProductCreditsOrderHoldingQueue("UserProfileKey")
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = drProductCreditsOrderHoldingQueue("CustomerKey")
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = "WEB_BOOKING_OVERDRFT"
        oCmdAddBooking.Parameters.Add(param2a)
        
        Dim param3 As SqlParameter = New SqlParameter("@BookingReference1", SqlDbType.NVarChar, 25)
        Dim param4 As SqlParameter = New SqlParameter("@BookingReference2", SqlDbType.NVarChar, 25)
        Dim param5 As SqlParameter = New SqlParameter("@BookingReference3", SqlDbType.NVarChar, 50)
        Dim param6 As SqlParameter = New SqlParameter("@BookingReference4", SqlDbType.NVarChar, 50)

        param3.Value = drProductCreditsOrderHoldingQueue("BookingReference1")
        param4.Value = drProductCreditsOrderHoldingQueue("BookingReference2")
        param5.Value = drProductCreditsOrderHoldingQueue("BookingReference3")
        param6.Value = drProductCreditsOrderHoldingQueue("BookingReference4")
        
        oCmdAddBooking.Parameters.Add(param3)
        oCmdAddBooking.Parameters.Add(param4)
        oCmdAddBooking.Parameters.Add(param5)
        oCmdAddBooking.Parameters.Add(param6)

        Dim param6a As SqlParameter = New SqlParameter("@ExternalReference", SqlDbType.NVarChar, 50)
        param6a.Value = drProductCreditsOrderHoldingQueue("ExternalReference")
        oCmdAddBooking.Parameters.Add(param6a)
        
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = drProductCreditsOrderHoldingQueue("SpecialInstructions")
        oCmdAddBooking.Parameters.Add(param7)
        
        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = drProductCreditsOrderHoldingQueue("PackingNoteInfo")
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
        param13.Value = drProductCreditsOrderHoldingQueue("CnorName")
        oCmdAddBooking.Parameters.Add(param13)
        
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = drProductCreditsOrderHoldingQueue("CnorAddr1")
        oCmdAddBooking.Parameters.Add(param14)
        
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = drProductCreditsOrderHoldingQueue("CnorAddr2")
        oCmdAddBooking.Parameters.Add(param15)
        
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = drProductCreditsOrderHoldingQueue("CnorAddr3")
        oCmdAddBooking.Parameters.Add(param16)
        
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = drProductCreditsOrderHoldingQueue("CnorTown")
        oCmdAddBooking.Parameters.Add(param17)
        
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = drProductCreditsOrderHoldingQueue("CnorState")
        oCmdAddBooking.Parameters.Add(param18)
        
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = drProductCreditsOrderHoldingQueue("CnorPostCode")
        oCmdAddBooking.Parameters.Add(param19)
        
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = drProductCreditsOrderHoldingQueue("CnorCountryKey")
        oCmdAddBooking.Parameters.Add(param20)
        
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = drProductCreditsOrderHoldingQueue("CnorCtcName")
        oCmdAddBooking.Parameters.Add(param21)
        
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = drProductCreditsOrderHoldingQueue("CnorTel")
        oCmdAddBooking.Parameters.Add(param22)
        
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = drProductCreditsOrderHoldingQueue("CnorEmail")
        oCmdAddBooking.Parameters.Add(param23)
        
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = drProductCreditsOrderHoldingQueue("CneeName")
        oCmdAddBooking.Parameters.Add(param25)

        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = drProductCreditsOrderHoldingQueue("CneeAddr1")
        oCmdAddBooking.Parameters.Add(param26)
        
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = drProductCreditsOrderHoldingQueue("CneeAddr2")
        oCmdAddBooking.Parameters.Add(param27)

        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = drProductCreditsOrderHoldingQueue("CneeAddr3")
        oCmdAddBooking.Parameters.Add(param28)
        
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = drProductCreditsOrderHoldingQueue("CneeTown")
        oCmdAddBooking.Parameters.Add(param29)
        
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = drProductCreditsOrderHoldingQueue("CneeState")
        oCmdAddBooking.Parameters.Add(param30)
        
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = drProductCreditsOrderHoldingQueue("CneePostCode")
        oCmdAddBooking.Parameters.Add(param31)

        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = drProductCreditsOrderHoldingQueue("CneeCountryKey")
        oCmdAddBooking.Parameters.Add(param32)
        
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = drProductCreditsOrderHoldingQueue("CneeCtcName")
        oCmdAddBooking.Parameters.Add(param33)

        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = drProductCreditsOrderHoldingQueue("CneeTel")
        oCmdAddBooking.Parameters.Add(param34)
        
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = drProductCreditsOrderHoldingQueue("CneeEmail")
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
        
        'For i As Int32 = 0 To oCmdAddBooking.Parameters.Count - 1
        '    Trace.Write(oCmdAddBooking.Parameters(i).ParameterName.ToString)
        '    Trace.Write(oCmdAddBooking.Parameters(i).DbType.ToString)
        '    If Not IsNothing(oCmdAddBooking.Parameters(i).Value) Then
        '        Trace.Write(oCmdAddBooking.Parameters(i).Value.ToString)
        '    Else
        '        Trace.Write("NOTHING")
        '    End If
        'Next
        
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
                For Each gvr As GridViewRow In gvOverdraftRequest.Rows
                    If gvr.RowType = DataControlRowType.DataRow Then
                        Dim hidLogisticProductKey As HiddenField = gvr.Cells(3).FindControl("hidLogisticProductKey")
                        Dim tbQty As TextBox = gvr.Cells(3).FindControl("tbQty")
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                        param51.Value = drProductCreditsOrderHoldingQueue("UserProfileKey")
                        oCmdAddStockItem.Parameters.Add(param51)
                        Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                        param52.Value = drProductCreditsOrderHoldingQueue("CustomerKey")
                        oCmdAddStockItem.Parameters.Add(param52)
                        Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                        param53.Value = lBookingKey
                        oCmdAddStockItem.Parameters.Add(param53)
                        Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param54.Value = CInt(hidLogisticProductKey.Value)
                        oCmdAddStockItem.Parameters.Add(param54)
                        Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                        param55.Value = "PENDING"
                        oCmdAddStockItem.Parameters.Add(param55)
                        Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                        param56.Value = CInt(tbQty.Text)
                        oCmdAddStockItem.Parameters.Add(param56)
                        Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
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

    Protected Sub btnEditDeliveryDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbName.Text = hidName.Value
        tbAddr1.Text = hidAddr1.Value
        tbAddr2.Text = hidAddr2.Value
        tbTown.Text = hidTown.Value
        tbPostcode.Text = hidPostcode.Value
        tbSpecialInstructions.Text = hidSpecialInstructions.Value
        pnlEditDeliveryDetails.Visible = True
    End Sub

    Protected Sub btnDeclineRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nProductCreditsOrderHoldingQueueKey As Int32 = b.CommandArgument
        Dim sSQL As String = "UPDATE ProductCreditsOrderHoldingQueue SET OrderStatus = 'DECLINED', MsgToOrderer = '" & tbMessageToOrderer.Text.Trim.Replace(Environment.NewLine, " ").Replace("'", "''") & "', OrderPlacedDateTime = GETDATE() WHERE [id] = " & nProductCreditsOrderHoldingQueueKey
        Call ExecuteQueryToDataTable(sSQL)
        Call HideAllPanelsAndRows()
        Call BindOverdraftRequests()
    End Sub
    
    Protected Sub btnSaveDeliveryDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        hidName.Value = tbName.Text
        hidAddr1.Value = tbAddr1.Text
        hidAddr2.Value = tbAddr2.Text
        hidTown.Value = tbTown.Text
        hidPostcode.Value = tbPostcode.Text
        hidSpecialInstructions.Value = tbSpecialInstructions.Text
        Call BuildDeliveryString()
        pnlEditDeliveryDetails.Visible = False
    End Sub

    Protected Sub btnCancelDeliveryDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlEditDeliveryDetails.Visible = False
    End Sub
    
    Protected Sub BuildDeliveryString()
        lblDelivery.Text = hidName.Value & " " & hidAddr1.Value & " " & hidAddr2.Value & " " & hidTown.Value & " " & hidPostcode.Value & " " & hidSpecialInstructions.Value
    End Sub
    
    Protected Sub lnkbtnCancelProcessingRequest_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanelsAndRows()
        pnlOverdraftRequests.Visible = True
    End Sub
    
    Protected Sub cbIncludeDeclinedRequests_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindOverdraftRequests()
    End Sub

    Protected Sub cbIncludeAuthorisedRequests_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindOverdraftRequests()
    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindOverdraftRequests()
    End Sub
    
    Protected Sub gvOverdraftRequests_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hidOrderStatus As HiddenField = gvr.Cells(2).FindControl("hidOrderStatus")
            If hidOrderStatus.Value.ToLower = "authorised" Then
                gvr.BackColor = Drawing.Color.PaleGreen
            ElseIf hidOrderStatus.Value.ToLower = "declined" Then
                gvr.BackColor = Drawing.Color.OrangeRed
            End If
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Product Credit Overdrafts</title>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server" />
    <asp:Panel ID="pnlOverdraftRequests" runat="server" Visible="false" Width="100%">
    <table style="width: 100%">
        <tr>
            <td style="width: 2%">
                &nbsp;
            </td>
            <td style="width: 26%">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="Small" Font-Names="Verdana" Font-Bold="True" ForeColor="Gray">Product Credits</asp:Label>
            </td>
            <td style="width: 40%">
                &nbsp;
            </td>
            <td style="width: 30%">
                &nbsp;
            </td>
            <td style="width: 2%">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="3">
                &nbsp;<asp:Label ID="lblLegendLastNPDFUploads1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Agent overdraft requests, last</asp:Label>
                &nbsp;<asp:DropDownList ID="ddlRequestDays" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onselectedindexchanged="DropDownList1_SelectedIndexChanged" AutoPostBack="True">
                    <asp:ListItem Selected="True" Value="11">10</asp:ListItem>
                    <asp:ListItem Value="31">30</asp:ListItem>
                    <asp:ListItem Value="91">90</asp:ListItem>
                </asp:DropDownList>
                &nbsp;<asp:Label ID="lblLegendLastNPDFUploads5" runat="server" Font-Names="Verdana" Font-Size="XX-Small">days</asp:Label>
                &nbsp;<asp:CheckBox ID="cbIncludeDeclinedRequests" runat="server" Font-Names="Verdana" Font-Size="XX-Small" oncheckedchanged="cbIncludeDeclinedRequests_CheckedChanged" Text="include declined requests" AutoPostBack="True" />
                &nbsp;<asp:CheckBox ID="cbIncludeAuthorisedRequests" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="include authorised requests" oncheckedchanged="cbIncludeAuthorisedRequests_CheckedChanged" AutoPostBack="True" />
                <asp:GridView ID="gvOverdraftRequests" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False" OnRowDataBound="gvOverdraftRequests_RowDataBound">
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:Button ID="btnView" runat="server" CommandArgument='<%# DataBinder.Eval(Container.DataItem, "id") %>' OnClick="btnView_Click" Text="view" Width="60px" />
                            </ItemTemplate>
                            <ItemStyle HorizontalAlign="Center" Width="70px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="OrderCreatedDateTime" HeaderText="Raised On" ReadOnly="True" SortExpression="OrderCreatedDateTime" />
                        <asp:BoundField DataField="AgentName" HeaderText="Agent" ReadOnly="True" SortExpression="AgentName" />
                        <asp:TemplateField HeaderText="Summary">
                            <ItemTemplate>
                                <asp:Label ID="lblSummary" runat="server" Text='<%# sGetSummary(Container.DataItem) %>' />
                                <asp:HiddenField ID="hidOrderStatus" Value='<%# DataBinder.Eval(Container.DataItem, "OrderStatus") %>' runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EmptyDataTemplate>
                        (no unprocessed overdraft requests found)
                    </EmptyDataTemplate>
                </asp:GridView>
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
            <td>
                &nbsp;
            </td>
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
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
    </table>
    </asp:Panel>
    <asp:Panel ID="pnlOverdraftRequest" runat="server" Visible="false" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 2%">
                </td>
                <td style="width: 26%">
                </td>
                <td style="width: 40%">
                </td>
                <td style="width: 30%">
                </td>
                <td style="width: 2%">
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="left" colspan="3">
                    <asp:Label ID="lblLegendLastNPDFUploads3" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Request raised on:</asp:Label>
                    &nbsp;<asp:Label ID="lblRequestRaisedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True"></asp:Label>
                    &nbsp;<asp:Label ID="lblLegendLastNPDFUploads2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">by Agent:</asp:Label>
                    &nbsp;<asp:Label ID="lblAgent" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:Label ID="lblLegendLastNPDFUploads4" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Delivery to:</asp:Label>
                    &nbsp;<asp:Label ID="lblDelivery" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                    </td>
                <td>
                    &nbsp;
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td>
                    <asp:Label ID="lblLegendItemsRequested" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Item(s) requested:</asp:Label>
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvOverdraftRequest" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False">
                        <Columns>
                            <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductCode" />
                            <asp:BoundField DataField="ProductDate" HeaderText="Value / Date" ReadOnly="True" SortExpression="ProductDate" />
                            <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                            <asp:TemplateField HeaderText="Qty">
                                <ItemTemplate>
                                    <asp:TextBox ID="tbQty" runat="server" Width="70px" Text='<%# DataBinder.Eval(Container.DataItem, "ItemsOut") %>'/>
                                    <asp:HiddenField ID="hidLogisticProductKey" Value='<%# DataBinder.Eval(Container.DataItem, "LogisticProductKey") %>' runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
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
                    <asp:Label ID="lblLegendMessageToOrderer" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Message to orderer:</asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
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
                <td colspan="3">
                    <asp:TextBox ID="tbMessageToOrderer" runat="server" Width="100%" 
                        TextMode="MultiLine" Font-Names="Verdana" Font-Size="XX-Small"></asp:TextBox>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td colspan="2">
                    <asp:Button ID="btnAuthoriseRequest" runat="server" Text="authorise request" onclick="btnAuthoriseRequest_Click" />
                    &nbsp;
                    <asp:Button ID="btnEditDeliveryDetails" runat="server" Text="edit delivery details" onclick="btnEditDeliveryDetails_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnDeclineRequest" runat="server" Text="decline request" onclick="btnDeclineRequest_Click" />
                    &nbsp;&nbsp; &nbsp;<asp:HiddenField ID="hidContactName" runat="server" />
                    <asp:HiddenField ID="hidName" runat="server" />
                    <asp:HiddenField ID="hidAddr1" runat="server" />
                    <asp:HiddenField ID="hidAddr2" runat="server" />
                    <asp:HiddenField ID="hidTown" runat="server" />
                    <asp:HiddenField ID="hidPostcode" runat="server" />
                    <asp:HiddenField ID="hidSpecialInstructions" runat="server" />
                </td>
                <td align="right">
                    <asp:LinkButton ID="lnkbtnCancelProcessingRequest" runat="server" 
                        onclick="lnkbtnCancelProcessingRequest_Click">go back</asp:LinkButton>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlEditDeliveryDetails" runat="server" Visible="false" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 2%">
                </td>
                <td style="width: 26%">
                    <asp:Label ID="lblLegendItemsRequested0" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Edit delivery details:</asp:Label>
                </td>
                <td style="width: 40%">
                </td>
                <td style="width: 30%">
                </td>
                <td style="width: 2%">
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="lblLegendAddresseeContactName" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Contact:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbContactName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="50" Width="100%" />
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td align="right">
                    &nbsp;
                    <asp:Label ID="lblLegendAddressee" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Name:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbName" runat="server" Width="100%" MaxLength="50" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
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
                <td align="right">
                    &nbsp;
                    <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Addr 1:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbAddr1" runat="server" Width="100%" MaxLength="50" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
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
                <td align="right">
                    &nbsp;
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Addr 2:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbAddr2" runat="server" Width="100%" MaxLength="50" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
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
                <td align="right">
                    &nbsp;
                    <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Town / City:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbTown" runat="server" Width="100%" MaxLength="50" Font-Names="Verdana" Font-Size="XX-Small"/>
                </td>
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
                <td align="right">
                    &nbsp;
                    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Post code:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbPostcode" runat="server" Width="100%" MaxLength="50" Font-Names="Verdana" Font-Size="XX-Small" />
                </td>
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
                <td align="right">
                    &nbsp;
                    <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small">Special Instructions:</asp:Label>
                </td>
                <td>
                    <asp:TextBox ID="tbSpecialInstructions" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="1000" TextMode="MultiLine" Width="100%" />
                </td>
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
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnSaveDeliveryDetails" runat="server" onclick="btnSaveDeliveryDetails_Click" Text="save" Width="120px" />
                    &nbsp;<asp:Button ID="btnCancelDeliveryDetails" runat="server" onclick="btnCancelDeliveryDetails_Click" Text="cancel" Width="80px" />
                </td>
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
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
