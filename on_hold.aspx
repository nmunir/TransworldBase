<%@ Page Language="VB" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

    '   Track & Trace page for pre-alert hyperlinks. Users arrive here without logging into the system. Only present information concerning the consignment key within the requesting page's querystring.
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Dim sCustServicesTel As String = ConfigurationManager.AppSettings("Cust_Services_Tel")
            lblOnHoldMessage.Text = "This Stock Booking has NOT been processed and is currently 'on hold' because one " & _
                                    "or more items did not have sufficient stock available to complete the order. This " & _
                                    "can occur on a multi-user system when stock is taken by another user before your " & _
                                    "order has been completed. <br/><br/> Please make the necessary adjustments " & _
                                    "below before re-submitting your order. If you wish to cancel this order then " & _
                                    "contact Customer Services on " & sCustServicesTel & " for assistance."
    
            lConsignmentKey = CLng(Request.QueryString("key"))
            Call GetStockBookingHeader()
            If sStateId = "ON_HOLD" Then
                Call GetStockItems()
                Call ShowStockBooking()
            Else
                Call ShowMessage()
            End If
        End If
    End Sub
    
    Protected Sub ShowStockBooking()
        pnlConsignmentDetail.Visible = True
        pnlMessage.Visible = False
        pnlConfirmation.Visible = False
    End Sub
    
    Protected Sub ShowMessage()
        pnlConsignmentDetail.Visible = False
        pnlMessage.Visible = True
        pnlConfirmation.Visible = False
    End Sub
    
    Protected Sub ShowConfirmation()
        pnlConsignmentDetail.Visible = False
        pnlMessage.Visible = False
        pnlConfirmation.Visible = True
    End Sub
    
    Protected Sub btn_RemoveItem_click(ByVal sender As Object, ByVal e As ImageClickEventArgs)
        Call RemoveItem()
    End Sub
    
    Protected Sub RemoveItem()
        'Dim dgi As DataGridItem
        Dim cb As CheckBox
        For Each dgi As DataGridItem In grid_StockItems.Items
            cb = CType(dgi.Cells(1).Controls(1), CheckBox)
            If cb.Checked Then
                Call RemoveStockItemFromBooking(CLng(dgi.Cells(0).Text))
            End If
        Next dgi
        Call GetStockBookingHeader()
        If sStateId = "ON_HOLD" Then
            Call GetStockItems()
            Call ShowStockBooking()
        Else
            Call ShowMessage()
        End If
    End Sub
    
    Protected Sub btn_ReSubmit_click(s As Object, e As ImageClickEventArgs)
        Call Resubmit()
    End Sub
    
    Protected Sub Resubmit()
        If isValidPick() Then
            For Each dgi As DataGridItem In grid_StockItems.Items
                Dim tbPickQuantity As TextBox = CType(dgi.Cells(6).FindControl("txtPickQuantity"), TextBox)
                Dim sQtyToPick As String = tbPickQuantity.Text.ToString
                If Not CInt(tbPickQuantity.Text) = 0 Then
                    Call UpdateStockItem(CLng(dgi.Cells(0).Text), CLng(sQtyToPick))
                Else
                    Call RemoveStockItemFromBooking(CLng(dgi.Cells(0).Text))
                End If
            Next dgi
            Call ResubmitStockBooking()
        End If
    End Sub
    
    Protected Sub GetStockBookingHeader()
        If lConsignmentKey > 0 Then
            lblError.Text = ""
            Dim sbHTML1 As StringBuilder = New StringBuilder()
            Dim sbHTML2 As StringBuilder = New StringBuilder()
            Dim sCnorName As String = String.Empty
            Dim sCnorAddr1 As String = String.Empty
            Dim sCnorAddr2 As String = String.Empty
            Dim sCnorAddr3 As String = String.Empty
            Dim sCnorTownCounty As String = String.Empty
            Dim sCnorPostCodeCountry As String = String.Empty
            Dim sCnorContact As String = String.Empty
            Dim sCneeName = String.Empty
            Dim sCneeAddr1 = String.Empty
            Dim sCneeAddr2 = String.Empty
            Dim sCneeAddr3 = String.Empty
            Dim sCneeTownCounty = String.Empty
            Dim sCneePostCodeCountry = String.Empty
            Dim sCneeContact As String = String.Empty
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_StockBooking_GetOnHold", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@Key", SqlDbType.Int, 4)
            oCmd.Parameters.Add(oParam)
            oParam.Value = lConsignmentKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                If Not IsDBNull(oDataReader("StateId")) Then
                    sStateId = oDataReader("StateId")
                End If
                If Not IsDBNull(oDataReader("RunningHeaderImage")) Then
                    sRunningHeaderImage = oDataReader("RunningHeaderImage")
                End If
                If Not IsDBNull(oDataReader("CustomerAccountCode")) Then
                    lblAcctCode.Text = oDataReader("CustomerAccountCode")
                End If
                If Not IsDBNull(oDataReader("AWB")) Then
                    lblConsignment.Text = oDataReader("AWB")
                End If
                If Not IsDBNull(oDataReader("CreatedOn")) Then
                    lblBookedOn.Text = Format(oDataReader("CreatedOn"), "dd.MM.yy")
                End If
                If Not IsDBNull(oDataReader("CreatedBy")) Then
                    lblBookedBy.Text = oDataReader("CreatedBy")
                End If
                If Not IsDBNull(oDataReader("ShippingInformation")) Then
                    lblShippingInfo.Text = oDataReader("ShippingInformation")
                End If
                If Not IsDBNull(oDataReader("SpecialInstructions")) Then
                    lblSpclInstructions.Text = oDataReader("SpecialInstructions")
                End If
                If Not IsDBNull(oDataReader("CustomerRef1")) Then
                    lblCustRef1.Text = oDataReader("CustomerRef1")
                End If
                If Not IsDBNull(oDataReader("CustomerRef2")) Then
                    lblCustRef2.Text = oDataReader("CustomerRef2")
                End If
                If Not IsDBNull(oDataReader("Misc1")) Then
                    lblCustRef3.Text = oDataReader("Misc1")
                End If
                If Not IsDBNull(oDataReader("Misc2")) Then
                    lblCustRef4.Text = oDataReader("Misc2")
                End If
                If Not IsDBNull(oDataReader("CnorName")) Then
                    sCnorName = oDataReader("CnorName")
                End If
                If Not IsDBNull(oDataReader("CnorAddr1")) Then
                    sCnorAddr1 = oDataReader("CnorAddr1")
                End If
                If Not IsDBNull(oDataReader("CnorAddr2")) Then
                    sCnorAddr2 = oDataReader("CnorAddr2")
                End If
                If Not IsDBNull(oDataReader("CnorAddr3")) Then
                    sCnorAddr3 = oDataReader("CnorAddr3")
                End If
                If Not IsDBNull(oDataReader("CnorTown")) Then
                    sCnorTownCounty = oDataReader("CnorTown")
                End If
                If Not IsDBNull(oDataReader("CnorState")) Then
                    sCnorTownCounty &= "  " & oDataReader("CnorState")
                End If
                If Not IsDBNull(oDataReader("CnorPostCode")) Then
                    sCnorPostCodeCountry = oDataReader("CnorPostCode")
                End If
                If Not IsDBNull(oDataReader("CnorCountryName")) Then
                    sCnorPostCodeCountry &= "  " & oDataReader("CnorCountryName")
                End If
                If Not IsDBNull(oDataReader("CnorCtcName")) Then
                    sCnorContact = oDataReader("CnorCtcName")
                End If
                If Not IsDBNull(oDataReader("CnorTel")) Then
                    sCnorContact &= "  " & oDataReader("CnorTel")
                End If
    
                If Not IsDBNull(oDataReader("CneeName")) Then
                    sCneeName = oDataReader("CneeName")
                End If
                If Not IsDBNull(oDataReader("CneeAddr1")) Then
                    sCneeAddr1 = oDataReader("CneeAddr1")
                End If
                If Not IsDBNull(oDataReader("CneeAddr2")) Then
                    sCneeAddr2 = oDataReader("CneeAddr2")
                End If
                If Not IsDBNull(oDataReader("CneeAddr3")) Then
                    sCneeAddr3 = oDataReader("CneeAddr3")
                End If
                If Not IsDBNull(oDataReader("CneeTown")) Then
                    sCneeTownCounty = oDataReader("CneeTown")
                End If
                If Not IsDBNull(oDataReader("CneeState")) Then
                    sCneeTownCounty &= "  " & oDataReader("CneeState")
                End If
                If Not IsDBNull(oDataReader("CneePostCode")) Then
                    sCneePostCodeCountry = oDataReader("CneePostCode")
                End If
                If Not IsDBNull(oDataReader("CneeCountryName")) Then
                    sCneePostCodeCountry &= "  " & oDataReader("CneeCountryName")
                End If
                If Not IsDBNull(oDataReader("CneeCtcName")) Then
                    sCneeContact = oDataReader("CneeCtcName")
                End If
                If Not IsDBNull(oDataReader("CneeTel")) Then
                    sCneeContact &= "  " & oDataReader("CneeTel")
                End If
                oDataReader.Close()
            Catch ex As SqlException
                lblError.Text = ex.ToString
                'Server.Transfer("error.aspx")
            Finally
                'oDataReader.Close()
                oConn.Close()
            End Try
    
            sbHTML1.Append(sCnorName & "<br/>" & vbCrLf)
            sbHTML1.Append(sCnorAddr1 & "<br/>" & vbCrLf)
            If sCnorAddr2 <> "" Then
                sbHTML1.Append(sCnorAddr2 & "<br/>" & vbCrLf)
            End If
            If sCnorAddr3 <> "" Then
                sbHTML1.Append(sCnorAddr3 & "<br/>" & vbCrLf)
            End If
            sbHTML1.Append(sCnorTownCounty & "<br/>" & vbCrLf)
            sbHTML1.Append(sCnorPostCodeCountry & "<br/>" & vbCrLf)
            sbHTML1.Append(sCnorContact & "<br/>" & vbCrLf)
            lblCnor.Text = sbHTML1.ToString()
    
            sbHTML2.Append(sCneeName & "<br/>" & vbCrLf)
            sbHTML2.Append(sCneeAddr1 & "<br/>" & vbCrLf)
            If sCneeAddr2 <> "" Then
                sbHTML2.Append(sCneeAddr2 & "<br/>" & vbCrLf)
            End If
            If sCneeAddr3 <> "" Then
                sbHTML2.Append(sCneeAddr3 & "<br/>" & vbCrLf)
            End If
            sbHTML2.Append(sCneeTownCounty & "<br/>" & vbCrLf)
            sbHTML2.Append(sCneePostCodeCountry & "<br/>" & vbCrLf)
            sbHTML2.Append(sCneeContact & "<br/>" & vbCrLf)
            lblCnee.Text = sbHTML2.ToString()
    
            If sRunningHeaderImage = "default" Then
                HeaderImage.ImageUrl = ConfigurationManager.AppSettings("default_running_header_image")
            ElseIf Session("RunningHeaderImage") <> "" Then
                HeaderImage.ImageUrl = sRunningHeaderImage
            Else
                HeaderImage.Visible = False
            End If
        End If
    End Sub
    
    Protected Sub GetStockItems()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim sProc As String
        sProc = "spASPNET_Consignment_GetOnHoldItems"
        Dim oAdapter As New SqlDataAdapter(sProc, oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = lConsignmentKey
        Try
            oAdapter.Fill(oDataSet, "StockItems")
            Dim Source As DataView = oDataSet.Tables("StockItems").DefaultView
            If Source.Count > 0 Then
                grid_StockItems.DataSource = Source
                grid_StockItems.DataBind()
                grid_StockItems.Visible = True
            Else
                grid_StockItems.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
            'Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub RemoveStockItemFromBooking(lStockBookingDetailKey As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As New SqlCommand("spASPNET_Consignment_RemoveOnHoldStockItem", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCollection1 As SqlParameter = New SqlParameter("@LogisticMovementKey", SqlDbType.Int, 4)
        oParamCollection1.Value = lStockBookingDetailKey
        oCmd.Parameters.Add(oParamCollection1)
        Dim oParamCollection2 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        oParamCollection2.Value = lConsignmentKey
        oCmd.Parameters.Add(oParamCollection2)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "RemoveItem")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
        Catch ex As SqlException
            oTrans.Rollback("RemoveItem")
            lblError.Text = ex.ToString
            'Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub UpdateStockItem(lStockBookingDetailKey As Long, lQtyToPick As Long)
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As New SqlCommand("spASPNET_Consignment_UpdateOnHoldStockItem", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCollection1 As SqlParameter = New SqlParameter("@LogisticMovementKey", SqlDbType.Int, 4)
        oParamCollection1.Value = lStockBookingDetailKey
        oCmd.Parameters.Add(oParamCollection1)
        Dim oParamCollection2 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        oParamCollection2.Value = lConsignmentKey
        oCmd.Parameters.Add(oParamCollection2)
        Dim oParamCollection3 As SqlParameter = New SqlParameter("@QtyToPick", SqlDbType.Int, 4)
        oParamCollection3.Value = lQtyToPick
        oCmd.Parameters.Add(oParamCollection3)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "UpdateItem")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
        Catch ex As SqlException
            oTrans.Rollback("UpdateItem")
            lblError.Text = ex.ToString
            'Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function isValidPick() As Boolean
        Dim i As Integer
        Dim bValid As Boolean = True
        Dim bAtLeastOnePositivePickQuantity As Boolean = False
        lblPickError.Text = ""
        For Each dgi As DataGridItem In grid_StockItems.Items
            i = i + 1
            Dim tbPickQuantity As TextBox = CType(dgi.Cells(6).FindControl("txtPickQuantity"), TextBox)
            Dim sQtyAvailable As String = dgi.Cells(5).Text
            Dim sQtyToPick As String = tbPickQuantity.Text.ToString
            Dim sProductCode As String = dgi.Cells(2).Text
            If IsNumeric(sQtyToPick) And IsNumeric(sQtyAvailable) Then
                If (CLng(sQtyToPick) > CLng(sQtyAvailable)) Then
                    bValid = False
                    lblPickError.Text = "Row " & i & " (Product " & sProductCode & ") has a pick quantity that exceeds the available quantity"
                    Exit For
                ElseIf CLng(sQtyToPick) = 0 Then
                    'bValid = False
                    lblPickError.Text = "Row " & i & " has a pick quantity of zero. You can proceed with your order, but note that items with zero pick quantity will be removed from your order before it is placed"
                ElseIf CLng(sQtyToPick) < 0 Then
                    bValid = False
                    lblPickError.Text = "Row " & i & " (Product " & sProductCode & ") has a negative pick quantity - please remove item before proceeding"
                    Exit For
                Else
                    bAtLeastOnePositivePickQuantity = True
                End If
            Else
                bValid = False
                lblPickError.Text = "Please ensure all pick quantities have numeric values"
                Exit For
            End If
        Next dgi

        If Not bAtLeastOnePositivePickQuantity AndAlso bValid = True Then
            bValid = False
            If lblPickError.Text = String.Empty Then
                lblPickError.Text = "Order must have at least one item with a positive, non-zero, pick quantity"
            End If
        End If

        If bValid Then
            isValidPick = True
        Else
            isValidPick = False
        End If
    End Function
    
    Protected Sub ResubmitStockBooking()
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmd As New SqlCommand("spASPNET_Consignment_ReSubmitOnHold", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParamCollection1 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4)
        oParamCollection1.Value = lConsignmentKey
        oCmd.Parameters.Add(oParamCollection1)
        Try
            oConn.Open()
            oTrans = oConn.BeginTransaction(IsolationLevel.ReadCommitted, "Resubmit")
            oCmd.Connection = oConn
            oCmd.Transaction = oTrans
            oCmd.ExecuteNonQuery()
            oTrans.Commit()
        Catch ex As SqlException
            oTrans.Rollback("Resubmit")
            lblError.Text = ex.ToString
            'Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
        lConsignmentKey = -1
        ShowConfirmation()
    End Sub

    Property lConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("OH_ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("OH_ConsignmentKey") = Value
        End Set
    End Property
    
    Property sStateId() As String
        Get
            Dim o As Object = ViewState("OH_StateId")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("OH_StateId") = Value
        End Set
    End Property
    
    Property sRunningHeaderImage() As String
        Get
            Dim o As Object = ViewState("OH_RunningHeaderImage")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("OH_RunningHeaderImage") = Value
        End Set
    End Property
    
    Protected Sub grid_StockItems_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dg As DataGrid = sender
        Dim dgiea As DataGridItemEventArgs = e
        Dim dgi As DataGridItem = dgiea.Item
        If dgi.ItemType = ListItemType.Item Then
            Dim nQtyRequested As Int32 = CInt(dgi.Cells(4).Text.Replace(",", ""))
            Dim nQtyAvailable As Int32 = CInt(dgi.Cells(5).Text.Replace(",", ""))
            If nQtyRequested >= nQtyAvailable Then
                dgi.Cells(4).ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub
    
    Protected Sub btnRemove_Click(sender As Object, e As System.EventArgs)
        Call RemoveItem()
    End Sub

    Protected Sub btnResubmit_Click(sender As Object, e As System.EventArgs)
        Call Resubmit()
    End Sub

    Protected Sub btnSetQuantity_Click(sender As Object, e As System.EventArgs)
        Call SetQuantity()
    End Sub
    
    Protected Sub SetQuantity()
        For Each dgi As DataGridItem In grid_StockItems.Items
            Dim nQtyRequested As Int32 = CInt(dgi.Cells(4).Text.Replace(",", ""))
            Dim nQtyAvailable As Int32 = CInt(dgi.Cells(5).Text.Replace(",", ""))
            Dim tbPickQuantity As TextBox = CType(dgi.Cells(6).FindControl("txtPickQuantity"), TextBox)
            If nQtyRequested <= nQtyAvailable Then
                tbPickQuantity.Text = nQtyRequested.ToString
            Else
                tbPickQuantity.BackColor = Drawing.Color.Red
            End If
        Next dgi
    End Sub

</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html lang="en" xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
    <title>Booking Advisory</title>
    <style type="text/css">@import url( CS_Style.css );
</style>
</head>
<body>
    <form id="Form1" method="post" enctype="multipart/form-data" runat="server">
        <asp:Panel id="pnlConsignmentDetail" runat="server" visible="False">
            <asp:Table id="tabHolding1" runat="server" Width="100%">
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabRunningHeader" Width="700px" runat="server" CellPadding="0" CellSpacing="0">
                            <asp:TableRow>
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="690px" HorizontalAlign="Left">
                                    <br />
                                    <asp:Image id="HeaderImage" runat="server" ></asp:Image>
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabConsignmentInfo" runat="server" Width="700px">
                            <asp:TableRow>
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="60px" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblAcctCode" font-size="X-Small"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell HorizontalAlign="Right" Width="330px" wrap="False">
                                    <asp:Label runat="server" font-size="X-Small">Consignment No: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="280px" wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblConsignment" wrap="False" font-size="X-Small" forecolor="Red"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                                <asp:TableCell></asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                                <asp:TableCell HorizontalAlign="Right" wrap="False">
                                    <asp:Label runat="server" font-size="X-Small">Booked On: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblBookedOn" font-size="X-Small" forecolor="Red"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow>
                                <asp:TableCell></asp:TableCell>
                                <asp:TableCell HorizontalAlign="Left">
                                    <asp:Label runat="server" font-size="X-Small" forecolor="Red">ON_HOLD</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell HorizontalAlign="Right" wrap="False">
                                    <asp:Label runat="server" font-size="X-Small">Booked By: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblBookedBy" font-size="X-Small" forecolor="Red"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell ColumnSpan="6">
                                    <br />
                                    <hr />
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabOnHoldMessages" runat="server" Width="700px">
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="680px" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblOnHoldMessage" font-size="X-Small"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell ColumnSpan="3">
                                    <br />
                                    <hr />
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <br />
                        <asp:Table id="tabDeliveryDetails" runat="server" Width="700px">
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                                <asp:TableCell Width="60px" wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">From: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="280px" wrap="False" VerticalAlign="Top" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCnor"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="60px" wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">To: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="280px" wrap="False" VerticalAlign="Top" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCnee"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabCustomerInstructions" runat="server" Width="700px">
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                                <asp:TableCell Width="150px" wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Shipping Note: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="530px" wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblShippingInfo" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell wrap="False"></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Special Instructions: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblSpclInstructions" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell wrap="False"></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Customer Ref 1: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCustRef1" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell wrap="False"></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Customer Ref 2: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCustRef2" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell wrap="False"></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Customer Ref 3: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCustRef3" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell wrap="False"></asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" verticalalign="Top">Customer Ref 4: </asp:Label>
                                </asp:TableCell>
                                <asp:TableCell wrap="False" HorizontalAlign="Left">
                                    <asp:Label runat="server" id="lblCustRef4" verticalalign="Top"></asp:Label>
                                </asp:TableCell>
                                <asp:TableCell></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="Table1" Width="700px" runat="server" CellPadding="0" CellSpacing="0">
                            <asp:TableRow>
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="690px" HorizontalAlign="Left">
                                <asp:Button ID="btnSetQuantity"  onclick="btnSetQuantity_Click" runat="server" Text="Set Qty Where Available" CausesValidation ="false" />
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <br />
                        <asp:Table id="tabStockItems" runat="server" Width="700px">
                            <asp:TableRow>
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                                <asp:TableCell Width="680px">
                                    <asp:DataGrid id="grid_StockItems" runat="server" Width="680px" AutoGenerateColumns="False" GridLines="None" ShowFooter="True" OnItemDataBound="grid_StockItems_ItemDataBound">
                                        <FooterStyle wrap="False"></FooterStyle>
                                        <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                                        <ItemStyle backcolor="White"></ItemStyle>
                                        <Columns>
                                            <asp:BoundColumn Visible="False" DataField="LogisticMovementKey" ReadOnly="True"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="Remove Item">
                                                <HeaderStyle verticalalign="Bottom" horizontalalign="Left"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" horizontalalign="Left"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:CheckBox id="chkRemove" runat="server"></asp:CheckBox>
                                                </ItemTemplate>
                                                <FooterStyle></FooterStyle>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="ProductCode" HeaderText="Product Code">
                                                <HeaderStyle verticalalign="Bottom" horizontalalign="Left"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" horizontalalign="Left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ProductDescription" HeaderText="Product Description">
                                                <HeaderStyle verticalalign="Bottom" horizontalalign="Left"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" horizontalalign="Left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ItemsOut" HeaderText="Quantity Requested" DataFormatString="{0:#,##0}">
                                                <HeaderStyle horizontalalign="Right"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" wrap="False" horizontalalign="Right"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ItemsAvailable" HeaderText="Quantity Available" DataFormatString="{0:#,##0}">
                                                <HeaderStyle horizontalalign="Right"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" wrap="False" horizontalalign="Right"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn>
                                                <HeaderStyle horizontalalign="Right"></HeaderStyle>
                                                <ItemStyle verticalalign="Top" wrap="False" horizontalalign="Right"></ItemStyle>
                                                <HeaderTemplate>
                                                    Amended Quantity
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:RequiredFieldValidator id="RequiredFieldValidator1" Font-Bold="True" ForeColor="Red" Runat="Server" Text=">>>" ControlToValidate="txtPickQuantity"></asp:RequiredFieldValidator>
                                                    <asp:TextBox id="txtPickQuantity" Width="40px" Font-Names="Verdana" Font-Size="XX-Small" runat="server"/>
                                                </ItemTemplate>
                                                <FooterStyle horizontalalign="Center"></FooterStyle>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </asp:TableCell>
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabButtons" runat="server" Width="700px">
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="340px" horizontalalign="Left">
                                    <asp:ImageButton onclick="btn_RemoveItem_click" runat="server" ImageUrl="./images/btn_remove.gif" CausesValidation="False" Visible="false"></asp:ImageButton><asp:Button ID="btnRemove" onclick="btnRemove_Click" runat="server" Text="Remove" />
                                </asp:TableCell>
                                <asp:TableCell Width="340px" horizontalalign="Right">
                                    <asp:ImageButton onclick="btn_ReSubmit_click" runat="server" ImageUrl="./images/btn_resubmit.gif" Visible="false"></asp:ImageButton><asp:Button ID="btnResubmit" onclick="btnResubmit_Click" runat="server" Text="Re-submit" />
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell horizontalalign="Right" ColumnSpan="4">
                                    <asp:Label id="lblPickError" runat="server" forecolor="Red"></asp:Label>
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <asp:Table id="tabFooter" runat="server" Width="700px">
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell Width="10px"></asp:TableCell>
                                <asp:TableCell Width="680px"></asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                            <asp:TableRow VerticalAlign="Top">
                                <asp:TableCell ColumnSpan="3">
                                    <br />
                                    <hr />
                                </asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Panel id="pnlMessage" runat="server" visible="False">
            <asp:Table id="tabHolding2" runat="server" Width="100%">
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <br />
                        <br />
                        <br />
                        <br />
                        <asp:Table id="tabMessage" runat="server" Width="700px">
                            <asp:TableRow HorizontalAlign="Center">
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                                <asp:TableCell Width="680px" wrap="False">
                                    <asp:Label runat="server">This link is no longer valid.</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Panel id="pnlConfirmation" runat="server" visible="False">
            <asp:Table id="tabHolding3" runat="server" Width="100%">
                <asp:TableRow HorizontalAlign="Center">
                    <asp:TableCell>
                        <br />
                        <br />
                        <br />
                        <br />
                        <asp:Table id="tabConfirmation" runat="server" Width="700px">
                            <asp:TableRow HorizontalAlign="Center">
                                <asp:TableCell Width="10px" wrap="False"></asp:TableCell>
                                <asp:TableCell Width="680px" wrap="False">
                                    <asp:Label runat="server">Your Stock Booking has been re-submitted.</asp:Label>
                                </asp:TableCell>
                                <asp:TableCell Width="10px"></asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <asp:Label id="lblError" runat="server" font-names="Verdana" font-size="X-Small" forecolor="#00C000"></asp:Label>
    </form>
</body>
</html>
