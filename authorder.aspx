<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Drawing.Color" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Dim oCmd As SqlCommand
    Dim sGUID As String
    Dim oDataTable As New DataTable
    Dim drAuthData As DataRow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        sGUID = Request.QueryString("GUID")
        If psOrderStatus = String.Empty Then
            If sGUID <> String.Empty Then
                Call FetchAuthRequest()
                Call ProcessAuthRequest()
            End If
        End If
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub FetchAuthRequest()
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_StockBooking_AuthOrderGetByGUID", oConn)
        Try
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@AuthorisationGUID", SqlDbType.VarChar, 20))
            oAdapter.SelectCommand.Parameters("@AuthorisationGUID").Value = sGUID
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show(ex.ToString)
        End Try
    End Sub

    Protected Sub ProcessAuthRequest()
        If oDataTable.Rows.Count > 0 Then
            drAuthData = oDataTable.Rows(0)
            psOrderStatus = drAuthData("OrderStatus")
            Call ShowAuth()
            If psOrderStatus <> "QUEUED" Then
                pnlAuthorise.Visible = False
                pnlView.Visible = True
                If psOrderStatus = "COMPLETE" Then
                    WebMsgBox.Show("This authorisation request has been approved")
                Else
                    WebMsgBox.Show("This authorisation request has been declined")
                End If
            Else
                pnlAuthorise.Visible = True
                pnlView.Visible = False
            End If
            Call GetAuthorisationOrder(drAuthData("id"))
        Else
            lblMessage.Text = "Authorisation request not found"
            Call ShowMessage()
        End If
    End Sub
    
    Protected Sub GetAuthorisationOrder(ByVal nHoldingQueueKey As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDatatable As DataTable = GetAuthOrderDetails(nHoldingQueueKey)
        lblAuthOrderOrderedBy.Text = drAuthData.Item("FirstName") & " " & drAuthData.Item("LastName")
        lblAuthOrderPlacedOn.Text = drAuthData.Item("OrderCreatedDateTime")
        lblAuthMsgToAuthoriser.Text = drAuthData.Item("MsgToAuthoriser")
        lblAuthOrderConsignee.Text = drAuthData.Item("CneeName")
        lblAuthOrderAttnOf.Text = drAuthData.Item("CneeCtcName")
        lblAuthOrderAddr1.Text = drAuthData.Item("CneeAddr1")
        lblAuthOrderAddr2.Text = drAuthData.Item("CneeAddr2")
        lblAuthOrderAddr3.Text = drAuthData.Item("CneeAddr3")
        lblAuthOrderTown.Text = drAuthData.Item("CneeTown")
        lblAuthOrderState.Text = drAuthData.Item("CneeState")
        lblAuthOrderPostcode.Text = drAuthData.Item("CneePostCode")
        lblAuthOrderCountry.Text = drAuthData.Item("CountryName")
        hidHoldingQueueKey.Value = nHoldingQueueKey
        If psOrderStatus = "QUEUED" Then
            gvAuthOrderDetails.DataSource = oDatatable
            gvAuthOrderDetails.DataBind()
        Else
            gvAuthOrderDetailsView.DataSource = oDatatable
            gvAuthOrderDetailsView.DataBind()
        End If
    End Sub
    
    Protected Function GetAuthOrderDetails(ByVal nHoldingQueueKey) As DataTable
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
    
    Protected Sub OrderAuthorise()
        Dim sValidationResult As String = CheckValidOrder()
        If sValidationResult.Length > 0 Then
            WebMsgBox.Show(sValidationResult)
        Else
            Dim lConsignmentKey As Long = SubmitOrder()
            If lConsignmentKey > 0 Then
                Call UpdateHoldingQueueEntry("COMPLETE", lConsignmentKey)
                Call EmailOrderer(bSuccess:=True, lConsignmentKey:=lConsignmentKey)
                lblMessage.Text = "Authorisation granted"
                Call ShowMessage()
            End If
        End If
        psOrderStatus = String.Empty
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
        Call UpdateHoldingQueueEntry("DECLINED", 0)
        Call EmailOrderer(bSuccess:=False, lConsignmentKey:=0)
        psOrderStatus = String.Empty
        lblMessage.Text = "Authorisation declined"
        Call ShowMessage()
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
    
    Protected Function SubmitOrder() As Long
        SubmitOrder = 0
        Dim drOrderDetails As DataRow = GetOrderAuthorisationByKey(hidHoldingQueueKey.Value)
        Dim sSpecialInstr As String
        Dim lBookingKey As Long
        Dim lConsignmentKey As Long
        Dim BookingFailed As Boolean
        Dim drv As DataRowView
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        
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

                        oCmdUpdateAuthorisedQuantity.Transaction = oTrans
                        oCmdUpdateAuthorisedQuantity.Connection = oConn
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
                    WebMsgBox.Show("No stock items found for booking")
                End If
            Else
                BookingFailed = True
                WebMsgBox.Show("Error adding Web Booking [BookingKey=0]")
            End If
            If Not BookingFailed Then
                oTrans.Commit()
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            WebMsgBox.Show(ex.ToString)
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
        If Not BookingFailed Then
            SubmitOrder = lConsignmentKey
        End If
    End Function

    Protected Sub btnOrderAuthorise_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call OrderAuthorise()
    End Sub
    
    Protected Sub ShowAuth()
        pnlOrderAuth.Visible = True
        pnlMessage.Visible = False
    End Sub
    
    Protected Sub ShowMessage()
        pnlMessage.Visible = True
        pnlOrderAuth.Visible = False
    End Sub
    
    Property psOrderStatus() As String
        Get
            Dim o As Object = ViewState("AO_OrderStatus")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AO_OrderStatus") = Value
        End Set
    End Property
    
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
    
    Protected Function gvAuthOrderDetailsItemForeColor(ByVal DataItem As Object) As System.Drawing.Color
        gvAuthOrderDetailsItemForeColor = Black
        If Not IsDBNull(DataBinder.Eval(DataItem, "Authorised")) AndAlso DataBinder.Eval(DataItem, "Authorised") = "N" Then
            gvAuthOrderDetailsItemForeColor = Red
        End If
    End Function

</script>

<html xmlns=" http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Quick Authorise</title>
</head>
<body style="font-family: Verdana">
    <form id="form1" runat="server">
    <div>
        <strong>Quick Authorise</strong><br />
            <br />
        <asp:Panel ID="pnlOrderAuth" runat="server" Visible="False" Width="100%">
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
                        <asp:Label ID="Label26" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Order created on:"></asp:Label></td>
                    <td style="height: 18px">
                        <asp:Label ID="lblAuthOrderPlacedOn" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text=""></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Orderer's message:"></asp:Label></td>
                    <td>
                        <asp:Label ID="lblAuthMsgToAuthoriser" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            Text=""></asp:Label></td>
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
                    <td style="height: 20px">
                        <asp:Label ID="Label36" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Postcode:"></asp:Label>
                    </td>
                    <td style="height: 20px">
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
            <asp:Panel ID="pnlAuthorise" runat="server" Width="100%" Visible="false">
            <asp:GridView ID="gvAuthOrderDetails" runat="server"  Font-Names="Verdana" Font-Size="XX-Small"
                 AutoGenerateColumns="False" Width="100%" GridLines="None" >
                <Columns>
                    <asp:TemplateField HeaderText="Product Code" >
                        <ItemTemplate>
                            <asp:HiddenField ID="hidLogisticProductKey" runat="server" Value='<%# Container.DataItem("LogisticProductKey")%>' />
                            <asp:HiddenField ID="hidQtyAvailable" runat="server" Value='<%# Container.DataItem("QtyAvailable")%>' />
                            <asp:HiddenField ID="hidArchiveFlag" runat="server"  Value='<%# Container.DataItem("ArchiveFlag")%>'/>
                            <asp:HiddenField ID="hidDeletedFlag" runat="server"  Value='<%# Container.DataItem("DeletedFlag")%>'/>
                          <asp:Label ID="lblAuthProductCode" runat="server" Text='<%# Container.DataItem("ProductCode")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Product Date" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDate" runat="server" Text='<%# Container.DataItem("ProductDate")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Description" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDescription" runat="server" Text='<%# Container.DataItem("ProductDescription")%>' ForeColor='<%# gvAuthOrderDetailsItemForeColor(Container.DataItem) %>' />
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty" >
                        <ItemTemplate>
                            <asp:TextBox ID="tbAuthOrderQty" Width="50px" MaxLength="6" BackColor="lightYellow" runat="server" Font-Names="Verdana" Font-Size="xX-Small"  Text='<%# Container.DataItem("ItemsOut")%>'></asp:TextBox>
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
            <strong style="font-size: xx-small">
            <br />
                Message to orderer (optional):<br />
            </strong>
            <asp:TextBox ID="tbAuthOrderMessage" runat="server" BackColor="LightYellow" TextMode="MultiLine" Width="100%" Rows="3" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="500"/>
            <br /><br />
            <asp:Button ID="btnOrderAuthorise" runat="server" Text="grant authorisation" OnClick="btnOrderAuthorise_Click" />
            <asp:Button ID="btnOrderDecline" runat="server" Text="decline authorisation" OnClick="btnOrderDecline_Click" />
            <br /><br />
            <asp:Label ID="lblValidationError" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Red"/>
            </asp:Panel>
            <asp:Panel ID="pnlView" runat="server" Width="100%" Visible="false">
            <asp:GridView ID="gvAuthOrderDetailsView" runat="server"  Font-Names="Verdana" Font-Size="XX-Small"
                 AutoGenerateColumns="False" Width="100%" GridLines="None" >
                <Columns>
                    <asp:TemplateField HeaderText="Product Code" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductCodeView" runat="server" Text='<%# Container.DataItem("ProductCode")%>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Product Date" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDateView" runat="server" Text='<%# Container.DataItem("ProductDate")%>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Description" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductDescriptionView" runat="server" Text='<%# Container.DataItem("ProductDescription")%>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Qty" >
                        <ItemTemplate>
                          <asp:Label ID="lblAuthProductItemsOutAuthorisedView" runat="server" Text='<%# Container.DataItem("ItemsOutAuthorised")%>' ></asp:Label>
                          (<asp:Label ID="lblAuthProductItemsOutView" runat="server" Text='<%# Container.DataItem("ItemsOut")%>' />)
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <EmptyDataTemplate>
                    no items found in order
                </EmptyDataTemplate>
                <RowStyle BackColor="WhiteSmoke" />
                <AlternatingRowStyle BackColor="White" />
            </asp:GridView>
            </asp:Panel>
        </asp:Panel>

        <asp:Panel ID="pnlMessage" runat="server" Width="100%" Visible="True">
            <asp:Label ID="lblMessage" runat="server" Font-Names="Verdana" Font-Size="X-Small"></asp:Label><br />
            <br />
            &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
            <asp:Button ID="btnClose" runat="server" OnClientClick="javascript:window.close();" Text="close window" />
        </asp:Panel>
        <br />
        <asp:HiddenField ID="hidHoldingQueueKey" runat="server" />
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/authorder.aspx?GUID=f86df27f-15b0-46c4-b" Visible="False">Call myself</asp:HyperLink></div>
    </form>
</body>
</html>