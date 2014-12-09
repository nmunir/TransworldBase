<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim dt As DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            pnlConsignment.Visible = False
            tbConsignmentToClone.Focus
        End If
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
    
    Protected Function Annul(drItem As Object) As String
        If IsDBNull(drItem) Then
            Annul = String.Empty
        Else
            Annul = drItem
        End If
    End Function

    Protected Function LoadConsignment(sConsignmentNumber As String) As Boolean
        LoadConsignment = False
        dt = New DataTable
        
        dt.Columns.Add(New DataColumn("ProductKey", GetType(String)))
        dt.Columns.Add(New DataColumn("ProductCode", GetType(String)))
        dt.Columns.Add(New DataColumn("Description", GetType(String)))
        dt.Columns.Add(New DataColumn("QtyToPick", GetType(Int32)))
        dt.Columns.Add(New DataColumn("QtyAvailable", GetType(Int32)))

        Dim sSQL As String = "SELECT cust.CustomerAccountCode, c.* FROM Consignment c INNER JOIN Customer cust ON c.CustomerKey = cust.CustomerKey WHERE AWB = '" & tbConsignmentToClone.Text.Replace("'", "''") & "'"
        Dim dtConsignment As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtConsignment.Rows.Count > 0 Then
            LoadConsignment = True
            Dim drConsignment As DataRow = dtConsignment.Rows(0)
            Dim dtLogisticBooking As DataTable = ExecuteQueryToDataTable("SELECT BookingOrigin, ShippingInformation FROM LogisticBooking WHERE LogisticBookingKey = " & drConsignment("StockBookingKey"))
            Dim drLogisticBooking = dtLogisticBooking.Rows(0)
            lblBookingOrigin.Text = drLogisticBooking("Bookingorigin")
            lblCustomer.Text = drConsignment("CustomerAccountCode")
            lblCustomerKey.Text = drConsignment("CustomerKey")
            lblUserKey.Text = drConsignment("UserKey")
            lblCreatedOn.Text = drConsignment("CreatedOn")
            tbCustRef1.Text = Annul(drConsignment("CustomerRef1"))
            tbCustRef2.Text = Annul(drConsignment("CustomerRef2"))
            tbCustRef3.Text = Annul(drConsignment("Misc1"))
            tbCustRef4.Text = Annul(drConsignment("Misc2"))
            tbExternalRef.Text = Annul(drConsignment("ExternalSystemId"))
            tbSpecialInstructions.Text = Annul(drConsignment("SpecialInstructions"))
            tbPackingNote.Text = drLogisticBooking("ShippingInformation")
            tbConsignmentType.Text = Annul(drConsignment("TypeId"))
            tbDescription.Text = Annul(drConsignment("Description"))
            lblConsignorName.Text = Annul(drConsignment("CnorName"))
            lblConsignorAddress1.Text = Annul(drConsignment("CnorAddr1"))
            lblConsignorAddress2.Text = Annul(drConsignment("CnorAddr2"))
            lblConsignorAddress3.Text = Annul(drConsignment("CnorAddr3"))
            lblConsignorTown.Text = Annul(drConsignment("CnorTown"))
            lblConsignorState.Text = Annul(drConsignment("CnorState"))
            lblConsignorPostcode.Text = Annul(drConsignment("CnorPostCode"))
            lblConsignorCountryKey.Text = drConsignment("CnorCountryKey")
            lblConsignorContactName.Text = Annul(drConsignment("CnorCtcName"))
            lblConsignorTelephone.Text = Annul(drConsignment("CnorTel"))
            lblConsignorEmail.Text = Annul(drConsignment("CnorEmail"))
            
            tbConsigneeName.Text = Annul(drConsignment("CneeName"))
            tbConsigneeAddr1.Text = Annul(drConsignment("CneeAddr1"))
            tbConsigneeAddr2.Text = Annul(drConsignment("CneeAddr2"))
            tbConsigneeAddr3.Text = Annul(drConsignment("CneeAddr3"))
            tbConsigneeTown.Text = Annul(drConsignment("CneeTown"))
            tbConsigneeState.Text = Annul(drConsignment("CneeState"))
            tbConsigneePostcode.Text = Annul(drConsignment("CneePostCode"))
            lblConsigneeCountry.Text = drConsignment("CneeCountryKey")
            tbConsigneeContactName.Text = Annul(drConsignment("CneeCtcName"))
            tbConsigneeTelephone.Text = Annul(drConsignment("CneeTel"))
            tbConsigneeEmail.Text = Annul(drConsignment("CneeEmail"))
            pnConsignmentKey = drConsignment("key")
            Call BindProducts()
        End If
    End Function
    
    Protected Sub BindProducts()
        Dim nConsignmentNumber As Int32 = pnConsignmentKey
        Dim sSQL As String = "SELECT * FROM LogisticMovement WHERE ConsignmentKey = " & nConsignmentNumber
        Dim dtMovements As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim dr As DataRow
        For Each drMovement As DataRow In dtMovements.Rows
            Dim drProduct As DataRow = ExecuteQueryToDataTable("SELECT ProductCode, ProductDescription FROM LogisticProduct WHERE LogisticProductKey = " & drMovement("LogisticProductKey")).Rows(0)
            dr = dt.NewRow()
            dr("ProductKey") = drMovement("LogisticProductKey")
            dr("QtyToPick") = drMovement("ItemsOut")
            dr("ProductCode") = drProduct("ProductCode")
            dr("Description") = drProduct("ProductDescription")
            dr("QtyAvailable") = GetQuantityInStock(drMovement("LogisticProductKey"))
            dt.Rows.Add(dr)
        Next
        gvProducts.DataSource = dt
        gvProducts.DataBind()
        Session("ProductTable") = dt
    End Sub
    
    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        tbConsignmentToClone.Text = tbConsignmentToClone.Text.Trim
        If tbConsignmentToClone.Text <> String.Empty Then
            If LoadConsignment(tbConsignmentToClone.Text) Then
                pnlConsignment.Visible = True
            Else
                WebMsgBox.Show("Could not load consignment " & tbConsignmentToClone.Text)
            End If
        Else
            WebMsgBox.Show("Please enter a valid consignment number")
        End If
    End Sub
    
    Protected Function GetQuantityInStock(ByVal nLogisticProductKey As Int32) As Int32
        GetQuantityInStock = -1
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetQuantityInStock", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = nLogisticProductKey
        Try
            oAdapter.Fill(oDataTable)
            GetQuantityInStock = oDataTable.Rows(0).Item(0)
        Catch ex As Exception
            ' report problem
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim nLogisticProductKey As Int32 = b.CommandArgument
        dt = Session("ProductTable")
        For Each dr As DataRow In dt.Rows
            If dr.Item("ProductKey") = nLogisticProductKey Then
                dr.Delete()
                Exit For
            End If
        Next
        Session("ProductTable") = dt
        gvProducts.DataSource = dt
        gvProducts.DataBind()
    End Sub

    Protected Sub SubmitOrder()
        Dim nBookingKey As Int32
        Dim nConsignmentKey As Int32
        Dim BookingFailed As Boolean
        Dim drv As DataRowView
        Dim oConn As New SqlConnection(gsConn)
        Dim oTrans As SqlTransaction
        Dim oCmdAddBooking As SqlCommand = New SqlCommand("spASPNET_StockBooking_Add3", oConn)
        oCmdAddBooking.CommandType = CommandType.StoredProcedure
        Dim param1 As SqlParameter = New SqlParameter("@UserProfileKey", SqlDbType.Int, 4)
        param1.Value = CInt(lblUserKey.Text)
        oCmdAddBooking.Parameters.Add(param1)
        Dim param2 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
        param2.Value = CInt(lblCustomerKey.Text)
        oCmdAddBooking.Parameters.Add(param2)
        Dim param2a As SqlParameter = New SqlParameter("@BookingOrigin", SqlDbType.NVarChar, 20)
        param2a.Value = lblBookingOrigin.Text
        oCmdAddBooking.Parameters.Add(param2a)

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
        param6a.Value = Nothing
        oCmdAddBooking.Parameters.Add(param6a)
            
        Dim param7 As SqlParameter = New SqlParameter("@SpecialInstructions", SqlDbType.NVarChar, 1000)
        param7.Value = tbSpecialInstructions.Text
        oCmdAddBooking.Parameters.Add(param7)

        Dim param8 As SqlParameter = New SqlParameter("@PackingNoteInfo", SqlDbType.NVarChar, 1000)
        param8.Value = tbPackingNote.Text
        oCmdAddBooking.Parameters.Add(param8)

        Dim param9 As SqlParameter = New SqlParameter("@ConsignmentType", SqlDbType.NVarChar, 20)
        param9.Value = tbConsignmentType.Text
        oCmdAddBooking.Parameters.Add(param9)

        Dim param10 As SqlParameter = New SqlParameter("@ServiceLevelKey", SqlDbType.Int, 4)
        param10.Value = -1
        oCmdAddBooking.Parameters.Add(param10)
        
        Dim param11 As SqlParameter = New SqlParameter("@Description", SqlDbType.NVarChar, 250)
        param11.Value = tbDescription.Text
        oCmdAddBooking.Parameters.Add(param11)
        Dim param13 As SqlParameter = New SqlParameter("@CnorName", SqlDbType.NVarChar, 50)
        param13.Value = lblConsignorName.Text
        oCmdAddBooking.Parameters.Add(param13)
        Dim param14 As SqlParameter = New SqlParameter("@CnorAddr1", SqlDbType.NVarChar, 50)
        param14.Value = lblConsignorAddress1.Text
        oCmdAddBooking.Parameters.Add(param14)
        Dim param15 As SqlParameter = New SqlParameter("@CnorAddr2", SqlDbType.NVarChar, 50)
        param15.Value = lblConsignorAddress2.Text
        oCmdAddBooking.Parameters.Add(param15)
        Dim param16 As SqlParameter = New SqlParameter("@CnorAddr3", SqlDbType.NVarChar, 50)
        param16.Value = lblConsignorAddress3.Text
        oCmdAddBooking.Parameters.Add(param16)
        Dim param17 As SqlParameter = New SqlParameter("@CnorTown", SqlDbType.NVarChar, 50)
        param17.Value = lblConsignorTown.Text
        oCmdAddBooking.Parameters.Add(param17)
        Dim param18 As SqlParameter = New SqlParameter("@CnorState", SqlDbType.NVarChar, 50)
        param18.Value = lblConsignorState.Text
        oCmdAddBooking.Parameters.Add(param18)
        Dim param19 As SqlParameter = New SqlParameter("@CnorPostCode", SqlDbType.NVarChar, 50)
        param19.Value = lblConsignorPostcode.Text
        oCmdAddBooking.Parameters.Add(param19)
        Dim param20 As SqlParameter = New SqlParameter("@CnorCountryKey", SqlDbType.Int, 4)
        param20.Value = lblConsignorCountryKey.Text
        oCmdAddBooking.Parameters.Add(param20)
        Dim param21 As SqlParameter = New SqlParameter("@CnorCtcName", SqlDbType.NVarChar, 50)
        param21.Value = lblConsignorContactName.Text
        oCmdAddBooking.Parameters.Add(param21)
        Dim param22 As SqlParameter = New SqlParameter("@CnorTel", SqlDbType.NVarChar, 50)
        param22.Value = lblConsignorTelephone.Text
        oCmdAddBooking.Parameters.Add(param22)
        Dim param23 As SqlParameter = New SqlParameter("@CnorEmail", SqlDbType.NVarChar, 50)
        param23.Value = lblConsignorEmail.Text
        oCmdAddBooking.Parameters.Add(param23)
        Dim param24 As SqlParameter = New SqlParameter("@CnorPreAlertFlag", SqlDbType.Bit)
        param24.Value = 0
        oCmdAddBooking.Parameters.Add(param24)
        Dim param25 As SqlParameter = New SqlParameter("@CneeName", SqlDbType.NVarChar, 50)
        param25.Value = tbConsigneeName.Text
        oCmdAddBooking.Parameters.Add(param25)
        Dim param26 As SqlParameter = New SqlParameter("@CneeAddr1", SqlDbType.NVarChar, 50)
        param26.Value = tbConsigneeAddr1.Text
        oCmdAddBooking.Parameters.Add(param26)
        Dim param27 As SqlParameter = New SqlParameter("@CneeAddr2", SqlDbType.NVarChar, 50)
        param27.Value = tbConsigneeAddr2.Text
        oCmdAddBooking.Parameters.Add(param27)
        Dim param28 As SqlParameter = New SqlParameter("@CneeAddr3", SqlDbType.NVarChar, 50)
        param28.Value = tbConsigneeAddr3.Text
        oCmdAddBooking.Parameters.Add(param28)
        Dim param29 As SqlParameter = New SqlParameter("@CneeTown", SqlDbType.NVarChar, 50)
        param29.Value = tbConsigneeTown.Text
        oCmdAddBooking.Parameters.Add(param29)
        Dim param30 As SqlParameter = New SqlParameter("@CneeState", SqlDbType.NVarChar, 50)
        param30.Value = tbConsigneeState.Text
        oCmdAddBooking.Parameters.Add(param30)
        Dim param31 As SqlParameter = New SqlParameter("@CneePostCode", SqlDbType.NVarChar, 50)
        param31.Value = tbConsigneePostcode.Text
        oCmdAddBooking.Parameters.Add(param31)
        Dim param32 As SqlParameter = New SqlParameter("@CneeCountryKey", SqlDbType.Int, 4)
        param32.Value = lblConsigneeCountry.Text
        oCmdAddBooking.Parameters.Add(param32)
        Dim param33 As SqlParameter = New SqlParameter("@CneeCtcName", SqlDbType.NVarChar, 50)
        param33.Value = tbConsigneeContactName.Text
        oCmdAddBooking.Parameters.Add(param33)
        Dim param34 As SqlParameter = New SqlParameter("@CneeTel", SqlDbType.NVarChar, 50)
        param34.Value = tbConsigneeTelephone.Text
        oCmdAddBooking.Parameters.Add(param34)
        Dim param35 As SqlParameter = New SqlParameter("@CneeEmail", SqlDbType.NVarChar, 50)
        param35.Value = tbConsigneeEmail.Text
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
            nBookingKey = CLng(oCmdAddBooking.Parameters("@LogisticBookingKey").Value.ToString)
            nConsignmentKey = CLng(oCmdAddBooking.Parameters("@ConsignmentKey").Value.ToString)
            If nBookingKey > 0 Then
                dt = Session("ProductTable")
                Dim BasketView As New DataView
                If dt.Rows.Count > 0 Then
                    For Each drProduct As DataRow In dt.Rows
                        'Dim lProductKey As Long = CLng(drv("ProductKey"))
                        'Dim lPickQuantity As Long = CLng(drv("QtyToPick"))
                        Dim oCmdAddStockItem As SqlCommand = New SqlCommand("spASPNET_LogisticMovement_Add", oConn)
                        oCmdAddStockItem.CommandType = CommandType.StoredProcedure
                        Dim param51 As SqlParameter = New SqlParameter("@UserKey", SqlDbType.Int, 4)
                        param51.Value = CInt(lblUserKey.Text)
                        oCmdAddStockItem.Parameters.Add(param51)
                        Dim param52 As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int, 4)
                        param52.Value = CInt(lblCustomerKey.Text)
                        oCmdAddStockItem.Parameters.Add(param52)
                        Dim param53 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                        param53.Value = nBookingKey
                        oCmdAddStockItem.Parameters.Add(param53)
                        Dim param54 As SqlParameter = New SqlParameter("@LogisticProductKey", SqlDbType.Int, 4)
                        param54.Value = drProduct("ProductKey")
                        oCmdAddStockItem.Parameters.Add(param54)
                        Dim param55 As SqlParameter = New SqlParameter("@LogisticMovementStateId", SqlDbType.NVarChar, 20)
                        param55.Value = "PENDING"
                        oCmdAddStockItem.Parameters.Add(param55)
                        Dim param56 As SqlParameter = New SqlParameter("@ItemsOut", SqlDbType.Int, 4)
                        param56.Value = drProduct("QtyToPick")
                        oCmdAddStockItem.Parameters.Add(param56)
                        Dim param57 As SqlParameter = New SqlParameter("@ConsignmentKey", SqlDbType.Int, 8)
                        param57.Value = nConsignmentKey
                        oCmdAddStockItem.Parameters.Add(param57)
                        oCmdAddStockItem.Connection = oConn
                        oCmdAddStockItem.Transaction = oTrans
                        oCmdAddStockItem.ExecuteNonQuery()
                    Next
                    Dim oCmdCompleteBooking As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_Complete", oConn)
                    oCmdCompleteBooking.CommandType = CommandType.StoredProcedure
                    Dim param71 As SqlParameter = New SqlParameter("@LogisticBookingKey", SqlDbType.Int, 4)
                    param71.Value = nBookingKey
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
                WebMsgBox.Show("Error adding Web Booking [BookingKey=0].")
            End If
            If Not BookingFailed Then
                oTrans.Commit()
                lblConsignmentNumber.Text = nConsignmentKey.ToString
            Else
                oTrans.Rollback("AddBooking")
            End If
        Catch ex As SqlException
            WebMsgBox.Show(ex.ToString)
            oTrans.Rollback("AddBooking")
        Finally
            oConn.Close()
        End Try
    End Sub

    Property pnConsignmentKey() As Int32
        Get
            Dim o As Object = ViewState("CC_ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("CC_ConsignmentKey") = Value
        End Set
    End Property

    Protected Sub btnPlaceOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SubmitOrder()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Clone Consignment</title>
    <style type="text/css">
       BODY {
        font-family: Verdana;
        font-size: xx-small
       }     
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <main:Header ID="ctlHeader" runat="server"/>
    <asp:Label ID="Label4" runat="server" Text="This utility lets you clone an existing consignment and then modify the contents and re-submit it (as a new consignment)" />
    <br />
    <br />
    <asp:Label ID="Label5" runat="server" 
        Text="This is useful in particular for an ON HOLD consignment that you want to adjust. Note that it does *not* modify the original consignment, which must be CANCELLED if it is ON HOLD." />
    <br />
    <br />
    <br />
    <asp:Label ID="Label6" runat="server" 
        Text="Consignment to clone:" />
    &nbsp;<asp:TextBox ID="tbConsignmentToClone" runat="server"></asp:TextBox>
&nbsp;<asp:Button ID="btnGo" runat="server" Text="go" onclick="btnGo_Click" />
    <br />
    <asp:Panel ID="pnlConsignment" runat="server" Width="100%">
        <table style="width: 100%">
            <tr>
                <td style="width: 1%">
                </td>
                <td style="width: 11%">
                    &nbsp;</td>
                <td style="width: 32%">
                    &nbsp;
                </td>
                <td style="width: 23%">
                </td>
                <td style="width: 32%">
                </td>
                <td style="width: 1%">
                </td>
            </tr>
            <tr>
                <td>
                    </td>
                <td align="right">
                    <asp:Label ID="Label7" runat="server" Text="Customer:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblCustomer" runat="server" />
                    &nbsp;(<asp:Label ID="lblCustomerKey" runat="server" />
                    )</td>
                <td>
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label10" runat="server" Text="Original AWB created:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblCreatedOn" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label49" runat="server" Text="Created by:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblUserKey" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label50" runat="server" Text="Booking origin:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblBookingOrigin" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td colspan="3">
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label8" runat="server" Text="Cust Ref 1:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbCustRef1" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label11" runat="server" Text="Cust Ref 2:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbCustRef2" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="25"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label12" runat="server" Text="Cust Ref 3:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbCustRef3" runat="server" Font-Names="Arial" Font-Size="XX-Small" MaxLength="50" Width="100%"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label13" runat="server" Text="Cust Ref 4:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbCustRef4" runat="server" Font-Names="Arial" Font-Size="XX-Small" MaxLength="50" Width="100%"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td colspan="3">
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label14" runat="server" Text="External Ref:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbExternalRef" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label15" runat="server" Text="Special Instructions:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbSpecialInstructions" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="1000"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label16" runat="server" Text="Packing note:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbPackingNote" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label17" runat="server" Text="Consignment type:" />
                    </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsignmentType" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="100%" MaxLength="20"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label18" runat="server" Text="Description:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbDescription" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="100%" MaxLength="250"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    </td>
                <td align="right">
                </td>
                <td colspan="3">
                </td>
                <td>
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label19" runat="server" Text="Consignor name:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorName" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label20" runat="server" Text="Consignor Address 1:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorAddress1" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label21" runat="server" Text="Consignor Address 2:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorAddress2" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    </td>
                <td align="right">
                    <asp:Label ID="Label1" runat="server" Text="Consignor Address 3:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorAddress3" runat="server" />
                </td>
                <td>
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label31" runat="server" Text="Consignor Town:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorTown" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label32" runat="server" Text="Consignor County / State:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorState" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label2" runat="server" Text="Consignor Post Code:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorPostcode" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label9" runat="server" Text="Consignor Country:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorCountryKey" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label23" runat="server" Text="Consignor Contact Name:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorContactName" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label27" runat="server" Text="Consignor Telephone:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorTelephone" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label25" runat="server" Text="Consignor Email:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsignorEmail" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                </td>
                <td colspan="3">
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label3" runat="server" Text="Consignee name:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeName" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label22" runat="server" Text="Consignee Address 1:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeAddr1" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label26" runat="server" Text="Consignee Address 2:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeAddr2" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    </td>
                <td align="right">
                    <asp:Label ID="Label33" runat="server" Text="Consignee Address 3:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeAddr3" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label35" runat="server" Text="Consignee Town:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeTown" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label37" runat="server" Text="Consignee County / State:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeState" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label39" runat="server" Text="Consignee Post Code:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneePostcode" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    </td>
                <td align="right">
                    <asp:Label ID="Label41" runat="server" Text="Consignee Country:" />
                </td>
                <td colspan="3">
                    <asp:Label ID="lblConsigneeCountry" runat="server" />
                </td>
                <td>
                    </td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label43" runat="server" Text="Consignee Contact Name:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeContactName" runat="server" Font-Names="Arial" 
                        Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label45" runat="server" Text="Consignee Telephone:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeTelephone" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label47" runat="server" Text="Consignee Email:" />
                </td>
                <td colspan="3">
                    <asp:TextBox ID="tbConsigneeEmail" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="100%" MaxLength="50"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                </td>
                <td colspan="3">
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    <asp:Label ID="Label48" runat="server" Text="Products:" />
                </td>
                <td colspan="3">
                    <asp:GridView ID="gvProducts" runat="server" CellPadding="2" Font-Names="Arial" Font-Size="XX-Small" Width="100%">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnRemove" runat="server" CommandArgument='<%# Container.DataItem("ProductKey")%>' onclick="btnRemove_Click" Text="remove" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
                <td>
                    &nbsp;</td>
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
            <tr>
                <td>
                    &nbsp;</td>
                <td align="right">
                    &nbsp;</td>
                <td colspan="3">
                    <asp:Button ID="btnPlaceOrder" runat="server" Text="Place order" onclick="btnPlaceOrder_Click" />
                    &nbsp;&nbsp;
                    <asp:Label ID="lblConsignmentNumber" runat="server" Font-Bold="True" Font-Size="Medium"></asp:Label>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
        <br />

    </asp:Panel>

    </form>
</body>
</html>