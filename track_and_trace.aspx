<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            ShowMainPanel()
            psFilterName = "MyLast20"
            BindConsignmentGrid("CreatedOn DESC")
            If Session("UserType") = "User" Then
                btn_Last50AllUsers.Visible = False
            End If
        End If
        Call SetTitle()
        Call SetStyleSheet()
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Track & Trace"
    End Sub
    
    Protected Sub SetStyleSheet()
        Dim hlCSSLink As New HtmlLink
        hlCSSLink.Href = Session("StyleSheetPath")
        hlCSSLink.Attributes.Add("rel", "stylesheet")
        hlCSSLink.Attributes.Add("type", "text/css")
        Page.Header.Controls.Add(hlCSSLink)
    End Sub

    Protected Sub HideAllPanels()
        pnlTrackAndTraceMain.Visible = False : pnlConsignmentDetail.Visible = False : pnlBookingDetail.Visible = False
    End Sub
    
    Protected Sub ShowMainPanel()
        Call HideAllPanels() : pnlTrackAndTraceMain.Visible = True
    End Sub
    
    Protected Sub ShowConsignmentDetail()
        Call HideAllPanels() : pnlConsignmentDetail.Visible = True
    End Sub
    
    Protected Sub ShowBookingDetail()
        pnlBookingDetail.Visible = True
    End Sub
    
    Protected Sub btn_MyLast20_click(ByVal s As Object, ByVal e As System.EventArgs)
        grid_Consignments.CurrentPageIndex = 0
        psFilterName = "MyLast20"
        BindConsignmentGrid("CreatedOn DESC")
        Call ShowMainPanel()
    End Sub
    
    Protected Sub btn_Last50AllUsers_Click(ByVal s As Object, ByVal e As System.EventArgs)
        grid_Consignments.CurrentPageIndex = 0
        psFilterName = "Last50AllUsers"
        BindConsignmentGrid("CreatedOn DESC")
        Call ShowMainPanel()
    End Sub
    
    Protected Sub btn_SearchAllConsignments_Click(ByVal s As Object, ByVal e As System.EventArgs)
        grid_Consignments.CurrentPageIndex = 0
        psFilterName = "SearchForAWB"
        BindConsignmentGrid("CreatedOn DESC")
        Call ShowMainPanel()
    End Sub
    
    Protected Sub btnByDate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sFromDate As String, sToDate As String
        Dim bValidDate As Boolean = True
        sFromDate = tbFromDate.Text.Trim()
        sToDate = tbToDate.Text.Trim()
        If Not (sFromDate = "" Or sToDate = "") Then
            Try
                sFromDate = DateTime.Parse(sFromDate)
                sToDate = DateTime.Parse(sToDate)
            Catch ex As Exception
                WebMsgBox.Show("Invalid date format - please check dates")
                bValidDate = False
                Exit Sub
            End Try
        Else
            WebMsgBox.Show("Please enter From and To dates")
            bValidDate = False
        End If

        If bValidDate Then
            grid_Consignments.CurrentPageIndex = 0
            psFilterName = "SearchByDate"
            BindConsignmentGrid("CreatedOn DESC")
            Call ShowMainPanel()
        End If
    End Sub

    Protected Sub btnOutstandingPODs_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        grid_Consignments.CurrentPageIndex = 0
        psFilterName = "OutstandingPODs"
        BindConsignmentGrid("CreatedOn DESC")
        Call ShowMainPanel()
    End Sub

    Protected Sub btn_RefreshTrackAndTraceGrid_click(ByVal s As Object, ByVal e As EventArgs)
        grid_Consignments.CurrentPageIndex = 0
        BindConsignmentGrid("CreatedOn DESC")
    End Sub
    
    Protected Sub btnBackToList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ShowMainPanel()
    End Sub

    Protected Sub grd_Consignments_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "tracking" Then
            Dim cell_Consignment As TableCell = e.Item.Cells(0)
            If IsNumeric(cell_Consignment.Text) Then
                plConsignmentKey = CLng(cell_Consignment.Text)
            End If
            GetConsignmentFromKey()
            ShowConsignmentDetail()
            If BindStockItems("ProductCode") > 0 Then Call ShowBookingDetail()
        End If
    End Sub
    
    Protected Sub grid_Consignments_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        grid_Consignments.CurrentPageIndex = e.NewPageIndex
        BindConsignmentGrid("CreatedOn DESC")
    End Sub
    
    Protected Sub BindConsignmentGrid(ByVal SortField As String)
        lblError.Text = ""
        lbl_AWBList.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter
        Select Case psFilterName
            Case "MyLast20"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_GetMyLast20", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserProfileKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@UserProfileKey").Value = Session("UserKey")
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
            Case "Last50AllUsers"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_TrackLast50All", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserProfileKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@UserProfileKey").Value = Session("UserKey")
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
            Case "SearchForAWB"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_SearchAll3", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = txtSearchAllConsignments.Text
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
                If Session("UserType") = "User" Then
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = CLng(Session("UserKey"))
                Else
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
                End If
            Case "SearchByDate"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_SearchByDate2", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = tbFromDate.Text
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = tbToDate.Text
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
                If Session("UserType") = "User" Then
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = CLng(Session("UserKey"))
                Else
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
                End If
            Case "OutstandingPODs"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Consignment_TrackOutstandingPODs2", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@UserKey", SqlDbType.Int))
                If Session("UserType") = "User" Then
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = CLng(Session("UserKey"))
                Else
                    oAdapter.SelectCommand.Parameters("@UserKey").Value = 0
                End If
        End Select
        Try
            oAdapter.Fill(oDataSet, "Consignments")
            Dim Source As DataView = oDataSet.Tables("Consignments").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                grid_Consignments.DataSource = Source
                grid_Consignments.DataBind()
                grid_Consignments.Visible = True
                btn_RefreshTrackAndTraceGrid.Visible = True
                If Source.Count > 10 Then
                    grid_Consignments.PagerStyle.Visible = True
                Else
                    grid_Consignments.PagerStyle.Visible = False
                End If
            Else
                grid_Consignments.Visible = False
                btn_RefreshTrackAndTraceGrid.Visible = False
                lbl_AWBList.Text = "no records found"
            End If
        Catch ex As SqlException
            Server.Transfer("error.aspx")
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub GetConsignmentFromKey()
        If plConsignmentKey > 0 Then
            ResetForm()
            lblError.Text = ""
            Dim sCnorName As String = String.Empty : Dim sCnorAddr1 = String.Empty : Dim sCnorAddr2 = String.Empty : Dim sCnorAddr3 As String = String.Empty : Dim sCnorTownCounty As String = String.Empty : Dim sCnorPostCodeCountry As String = String.Empty, sCnorContact As String = String.Empty
            Dim sCneeName As String = String.Empty : Dim sCneeAddr1 As String = String.Empty : Dim sCneeAddr2 As String = String.Empty : Dim sCneeAddr3 As String = String.Empty : Dim sCneeTownCounty As String = String.Empty : Dim sCneePostCodeCountry As String = String.Empty : Dim sCneeContact As String = String.Empty
            Dim oDataReader As SqlDataReader
            Dim oConn As New SqlConnection(gsConn)
            Dim oCmd As New SqlCommand("spASPNET_Tracking_GetConsignmentFromKey2", oConn)
            oCmd.CommandType = CommandType.StoredProcedure
            Dim oParam As New SqlParameter("@Key", SqlDbType.Int, 4)
            oCmd.Parameters.Add(oParam)
            oParam.Value = plConsignmentKey
            Try
                oConn.Open()
                oDataReader = oCmd.ExecuteReader()
                oDataReader.Read()
                lblConsignment.Text = plConsignmentKey
                If Not IsDBNull(oDataReader("CreatedOn")) Then lblDate.Text = Format(oDataReader("CreatedOn"), "dd.MM.yy")
                If Not IsDBNull(oDataReader("CnorName")) Then sCnorName = oDataReader("CnorName")
                If Not IsDBNull(oDataReader("CustomerAccountCode")) Then sCnorName &= " [" & oDataReader("CustomerAccountCode") & "]"
                If Not IsDBNull(oDataReader("CnorAddr1")) Then sCnorAddr1 = oDataReader("CnorAddr1")
                If Not IsDBNull(oDataReader("CnorAddr2")) Then sCnorAddr2 = oDataReader("CnorAddr2")
                If Not IsDBNull(oDataReader("CnorAddr3")) Then sCnorAddr3 = oDataReader("CnorAddr3")
                If Not IsDBNull(oDataReader("CnorTown")) Then sCnorTownCounty = oDataReader("CnorTown")
                If Not IsDBNull(oDataReader("CnorState")) Then sCnorTownCounty &= "  " & oDataReader("CnorState")
                If Not IsDBNull(oDataReader("CnorPostCode")) Then sCnorPostCodeCountry = oDataReader("CnorPostCode")
                If Not IsDBNull(oDataReader("CnorCountryName")) Then sCnorPostCodeCountry &= "  " & oDataReader("CnorCountryName")
                If Not IsDBNull(oDataReader("CnorCtcName")) Then sCnorContact = oDataReader("CnorCtcName")
                If Not IsDBNull(oDataReader("CnorTel")) Then sCnorContact &= "  " & oDataReader("CnorTel")
                If Not IsDBNull(oDataReader("CneeName")) Then sCneeName = oDataReader("CneeName")
                If Not IsDBNull(oDataReader("CneeAddr1")) Then sCneeAddr1 = oDataReader("CneeAddr1")
                If Not IsDBNull(oDataReader("CneeAddr2")) Then sCneeAddr2 = oDataReader("CneeAddr2")
                If Not IsDBNull(oDataReader("CneeAddr3")) Then sCneeAddr3 = oDataReader("CneeAddr3")
                If Not IsDBNull(oDataReader("CneeTown")) Then sCneeTownCounty = oDataReader("CneeTown")
                If Not IsDBNull(oDataReader("CneeState")) Then sCneeTownCounty &= "  " & oDataReader("CneeState")
                If Not IsDBNull(oDataReader("CneePostCode")) Then sCneePostCodeCountry = oDataReader("CneePostCode")
                If Not IsDBNull(oDataReader("CneeCountryName")) Then sCneePostCodeCountry &= "  " & oDataReader("CneeCountryName")
                If Not IsDBNull(oDataReader("CneeCtcName")) Then sCneeContact = oDataReader("CneeCtcName")
                If Not IsDBNull(oDataReader("CneeTel")) Then sCneeContact &= "  " & oDataReader("CneeTel")
                If Not IsDBNull(oDataReader("Weight")) AndAlso oDataReader("Weight") <> "0" Then lblWeight.Text = oDataReader("Weight")
                If Not IsDBNull(oDataReader("NOP")) AndAlso oDataReader("NOP") <> "0" Then lblNOP.Text = oDataReader("NOP")
                If Not IsDBNull(oDataReader("SpecialInstructions")) Then lblSpclInstructions.Text = oDataReader("SpecialInstructions")
                If Not IsDBNull(oDataReader("ShippingInformation")) Then lblPackingNote.Text = oDataReader("ShippingInformation")
                If Not IsDBNull(oDataReader("Description")) Then lblContents.Text = oDataReader("Description")
                If Not IsDBNull(oDataReader("ValForCustoms")) AndAlso oDataReader("ValForCustoms") > 0 Then lblCustomsValue.Text = oDataReader("ValForCustoms")
                If Not IsDBNull(oDataReader("CustomerRef1")) Then lblCustRef1.Text = oDataReader("CustomerRef1")
                If Not IsDBNull(oDataReader("CustomerRef2")) Then lblCustRef2.Text = oDataReader("CustomerRef2")
                If Not IsDBNull(oDataReader("Misc1")) Then lblCustRef3.Text = oDataReader("Misc1")
                If Not IsDBNull(oDataReader("Misc2")) Then lblCustRef4.Text = oDataReader("Misc2")
                If Not IsDBNull(oDataReader("PODDate")) Then lblPODDate.Text = oDataReader("PODDate")
                If Not IsDBNull(oDataReader("PODName")) Then lblPODName.Text = oDataReader("PODName")
                If Not IsDBNull(oDataReader("PODTime")) Then lblPODTime.Text = oDataReader("PODTime")
                oDataReader.Close()
            Catch ex As SqlException
                Server.Transfer("error.aspx")
            Finally
                oConn.Close()
            End Try
            lblCnorAddr1.Text = sCnorName : lblCnorAddr2.Text = sCnorAddr1 : lblCnorAddr3.Text = sCnorAddr2 : lblCnorAddr4.Text = sCnorAddr3 : lblCnorAddr5.Text = sCnorTownCounty : lblCnorAddr6.Text = sCnorPostCodeCountry : lblCnorAddr7.Text = sCnorContact
            lblCneeAddr1.Text = sCneeName : lblCneeAddr2.Text = sCneeAddr1 : lblCneeAddr3.Text = sCneeAddr2 : lblCneeAddr4.Text = sCneeAddr3 : lblCneeAddr5.Text = sCneeTownCounty : lblCneeAddr6.Text = sCneePostCodeCountry : lblCneeAddr7.Text = sCneeContact
            GetTracking()
        End If
    End Sub
    
    Protected Sub GetTracking()
        If plConsignmentKey > 0 Then
            lblError.Text = ""
            Dim oConn As New SqlConnection(gsConn)
            Dim oDataTable As New DataTable
            Dim oAdapter As New SqlDataAdapter("spASPNET_Consignment_GetTracking", oConn)
            Try
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int))
                oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = plConsignmentKey
                oAdapter.Fill(oDataTable)
                If oDataTable.Rows.Count > 0 Then
                    For Each dr As DataRow In oDataTable.Rows
                        If dr("Description").ToString.ToLower.Contains("with operations") Then
                            dr("Description") = "Job despatched"
                        End If
                    Next
                    grid_Tracking.DataSource = oDataTable
                    grid_Tracking.DataBind()
                    grid_Tracking.Visible = True
                Else
                    grid_Tracking.Visible = False
                End If
            Catch ex As SqlException
                lblError.Text = ex.ToString
            Finally
                oConn.Close()
            End Try
        End If
    End Sub
    
    Protected Function BindStockItems(ByVal SortField As String) As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_LogisticBooking_GetMovementsWithVals", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ConsignmentKey", SqlDbType.Int, 4))
        oAdapter.SelectCommand.Parameters("@ConsignmentKey").Value = plConsignmentKey
        Try
            oAdapter.Fill(oDataSet, "Movements")
            Dim Source As DataView = oDataSet.Tables("Movements").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                grd_BookingItems.DataSource = Source
                grd_BookingItems.DataBind()
                grd_BookingItems.Visible = True
            Else
                grd_BookingItems.Visible = False
            End If
            BindStockItems = CInt(Source.Count)
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub ResetForm()
        lblDate.Text = ""
        lblCnorAddr1.Text = "" : lblCnorAddr2.Text = "" : lblCnorAddr3.Text = "" : lblCnorAddr4.Text = "" : lblCnorAddr5.Text = ""
        lblCneeAddr1.Text = "" : lblCneeAddr2.Text = "" : lblCneeAddr3.Text = "" : lblCneeAddr4.Text = "" : lblCneeAddr5.Text = ""
        lblWeight.Text = "" : lblNOP.Text = "" : lblSpclInstructions.Text = "" : lblContents.Text = "" : lblCustRef1.Text = "" : lblCustRef2.Text = ""
        lblPODDate.Text = "" : lblPODName.Text = "" : lblPODTime.Text = ""
    End Sub

    Protected Sub cbMoreOptions_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If cbMoreOptions.Checked Then
            spnDateRange.Visible = True
            spnMainOptions.Visible = False
        Else
            spnMainOptions.Visible = True
            spnDateRange.Visible = False
        End If
    End Sub

    Property plConsignmentKey() As Long
        Get
            Dim o As Object = ViewState("TT_ConsignmentKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("TT_ConsignmentKey") = Value
        End Set
    End Property
    
    Property psFilterName() As String
        Get
            Dim o As Object = ViewState("TT_FilterName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("TT_FilterName") = Value
        End Set
    End Property
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Track &amp; Trace</title>
</head>
<body>
    <form id="Form1" method="post" enctype="multipart/form-data" runat="server">
        <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_tracktrace">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        <asp:Table id="tabTopButtons" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" style="font-family: Verdana">
            <asp:TableRow VerticalAlign="Middle" runat="server">
                <asp:TableCell HorizontalAlign="Left" VerticalAlign="Middle" wrap="False" runat="server"> 
                    <span id="spnMainOptions" runat="server">                       
                    <asp:Button ID="btn_MyLast20" runat="server" OnClick="btn_MyLast20_click" Text="my last 20 jobs" Tooltip="click here to see your last 20 consignments" />  
                    &nbsp;&nbsp;<asp:Button ID="btn_Last50AllUsers" runat="server" OnClick="btn_Last50AllUsers_Click" Text="last 50 all users" Tooltip="click here to see the last 50 consignments all users" />                       
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label16" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana">Search:</asp:Label> 
                    &nbsp;<asp:TextBox runat="server" Height="20px" Width="80px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtSearchAllConsignments" MaxLength="50"></asp:TextBox>                        
                    &nbsp;<asp:Button ID="btn_SearchAllConsignments" runat="server" OnClick="btn_SearchAllConsignments_Click" Text="go" Tooltip="wild card search across all consignments"/>
                    <a onmouseover="return escape('Searches for a consignment by Booking number, Air Waybill number, Customer Reference fields, Consignee name, Consignee address (Addr line 1, Town, County/State, Post Code, Country). You can enter a partial string, eg \'ger\' will match GERMANY.')" style="color:gray; cursor:help">&nbsp;?&nbsp;</a>
                    </span>
                    <span id="spnDateRange" runat="server" visible="False" style="font-size: xx-small; font-family: Verdana;">
                         From:
                        <asp:TextBox ID="tbFromDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90px"
                                     runat="server" />
                         <a href="javascript:;"
                            onclick="window.open('PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1"
                                 src="~/images/SmallCalendar.gif"
                                 runat="server"
                                 border="0"
                              IE:visible="True"
                                 visible="False"
                               ></a>
                               To:
                        <asp:TextBox ID="tbToDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90px"
                                     runat="server" />
                         <a href="javascript:;"
                            onclick="window.open('PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img2"
                                 src="~/images/SmallCalendar.gif"
                                 runat="server"
                                 border="0"
                              IE:visible="True"
                                 visible="False"
                               ></a>
                    &nbsp;&nbsp;<asp:Button ID="btnByDate" runat="server" Text="search by date" OnClick="btnByDate_Click" />                       
                    <a onmouseover="return escape('Click the calendar icon to select a date, or type the date directly in the format dd-mmm-yyyy, eg 29-Jan-2006 ')" style="color:gray; cursor:help">&nbsp;?&nbsp;</a>
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnOutstandingPODs" runat="server" Text="outstanding PODs" OnClick="btnOutstandingPODs_Click" />
                    </span>
                    &nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbMoreOptions" Text="more options" Font-Size="XX-Small" Font-Names="Verdana" runat="server" OnCheckedChanged="cbMoreOptions_CheckedChanged" AutoPostBack="True" />
                </asp:TableCell>
                <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False" runat="server"></asp:TableCell>
            </asp:TableRow>
        </asp:Table>
        
        <asp:Panel id="pnlTrackAndTraceMain" runat="server" visible="False" Width="100%">            
            <br />
            <asp:DataGrid id="grid_Consignments" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" OnItemCommand="grd_Consignments_item_click" ShowFooter="True" GridLines="None" AutoGenerateColumns="False" Visible="False" AllowPaging="True" OnPageIndexChanged="grid_Consignments_Page_Change">
                <FooterStyle wrap="False"></FooterStyle>
                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                <ItemStyle backcolor="LightGray"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="Key" visible="False">
                        <HeaderStyle></HeaderStyle>
                        <ItemStyle></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemStyle wrap="False" horizontalalign="Left"></ItemStyle>
                        <ItemTemplate>
                            <asp:Button ID="btnTracking" CommandName="tracking" runat="server" Text="tracking" ToolTip="view consignment details" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="AWB" HeaderText="Consignment">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CreatedOn" HeaderText="Created On" DataFormatString="{0:dd/MM/yy HH:mm}">
                        <HeaderStyle font-bold="True" wrap="False" horizontalalign="Left" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" HeaderText="Booked By">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeName" HeaderText="To">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" HeaderText="City">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="POD" HeaderText="P.O.D.">
                        <HeaderStyle font-bold="True" wrap="False" forecolor="Gray"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                        <ItemTemplate>
                            <asp:Button ID="btnPrintConsignment" OnClientClick='<%# "Javascript:TT_PrintConsignment(" & DataBinder.Eval(Container.DataItem,"Key") & ")" %> ' runat="server" Text="print" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
            </asp:DataGrid>
            &nbsp;<asp:Label id="lbl_AWBList" runat="server" font-names="Verdana" font-size="X-Small" forecolor="Blue"></asp:Label>
            <asp:LinkButton id="btn_RefreshTrackAndTraceGrid" onclick="btn_RefreshTrackAndTraceGrid_click" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Visible="False" CausesValidation="False" ForeColor="Blue">refresh</asp:LinkButton>
        </asp:Panel>
        
        <asp:Panel id="pnlConsignmentDetail" runat="server" visible="False" Width="100%">
            <asp:Table id="tabConsignmentDetail1" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small">
                <asp:TableRow>
                    <asp:TableCell VerticalAlign="Middle" Wrap="False" Width="450px">
                        <asp:Label runat="server" font-size="X-Small" font-names="Verdana" font-bold="True" forecolor="Gray">Consignment:</asp:Label> &nbsp;&nbsp;&nbsp;<asp:Label runat="server" id="lblConsignment" font-size="X-Small" forecolor="Red" font-names="Verdana"></asp:Label> &nbsp;&nbsp;&nbsp;
                        &nbsp;&nbsp;&nbsp;<asp:Label runat="server" font-size="X-Small" forecolor="Gray">Dated:</asp:Label> &nbsp;&nbsp;&nbsp;<asp:Label runat="server" id="lblDate" font-size="X-Small" forecolor="Red" font-names="Verdana"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False">
                        <asp:HyperLink runat="server" Visible="False" ImageUrl="./images/btn_help.gif" NavigateUrl="javascript:TT_OpenHelpWindow('./help/consignment_hlp.aspx');"></asp:HyperLink>
                    </asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right" VerticalAlign="Middle" wrap="False" Width="40px">
                        <asp:Button ID="btnBackToList" runat="server" Text="back to list" OnClick="btnBackToList_Click" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <br />
            <asp:Table id="tabConsignmentDetail2" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" forecolor="Navy">
                <asp:TableRow>
                    <asp:TableCell Width="15%"></asp:TableCell>
                    <asp:TableCell Width="35%"></asp:TableCell>
                    <asp:TableCell Width="10%"></asp:TableCell>
                    <asp:TableCell Width="40%"></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">From:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr1"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">To:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr1"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr2"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr2"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr3"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr3"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr4"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr4"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr5"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr5"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr6"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr6"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCnorAddr7"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCneeAddr7"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">NOP:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblNOP" ></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Weight:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblWeight" ></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Contents:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblContents"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Value:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCustomsValue"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Spcl Instr:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell ColumnSpan="3">
                        <asp:Label runat="server" id="lblSpclInstructions"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label ID="Label3" runat="server" Font-Bold="true" forecolor="Gray">Packing Note:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell ColumnSpan="3">
                        <asp:Label runat="server" id="lblPackingNote"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Ref 1:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCustRef1"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Ref 2:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCustRef2"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Ref 3:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCustRef3"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Ref 4:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell>
                        <asp:Label runat="server" id="lblCustRef4"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell>
                        <asp:Label runat="server" Font-Bold="true" forecolor="Gray">Received By:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell ColumnSpan="3">
                        <asp:Label runat="server" id="lblPODDate" forecolor="Red"></asp:Label> &nbsp;<asp:Label runat="server" id="lblPODName" forecolor="Red"></asp:Label> &nbsp;<asp:Label runat="server" id="lblPODTime" forecolor="Red"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <hr />
            <asp:table id="Table1" runat="Server" width="100%">
                <asp:TableRow>
                    <asp:TableCell Wrap="False">
                        <asp:Label ID="Label2" runat="server" Text="Tracking Events" forecolor="Gray" font-size="X-Small" font-names="Arial" font-bold="True" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:table>
            <asp:DataGrid id="grid_Tracking" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" GridLines="None" AutoGenerateColumns="False" Visible="False">
                <FooterStyle wrap="False"></FooterStyle>
                <HeaderStyle font-names="Verdana" wrap="False"></HeaderStyle>
                <PagerStyle font-size="X-Small" font-names="Verdana" font-bold="True" horizontalalign="Center" forecolor="Blue" backcolor="Silver" wrap="False" mode="NumericPages"></PagerStyle>
                <Columns>
                    <asp:BoundColumn DataField="Time" HeaderText="Time" DataFormatString="{0:dd.MM.yy HH:mm}">
                        <HeaderStyle wrap="False" Font-Bold="true" forecolor="Gray" width="15%"></HeaderStyle>
                        <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Location" HeaderText="Location">
                        <HeaderStyle wrap="False" Font-Bold="true" horizontalalign="Left" forecolor="Gray" width="10%"></HeaderStyle>
                        <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Description" HeaderText="Description">
                        <HeaderStyle wrap="False" Font-Bold="true" forecolor="Gray" width="75%"></HeaderStyle>
                        <ItemStyle forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <br />
        <asp:Panel id="pnlBookingDetail" runat="server" Visible="false">
             <hr />
            <asp:table id="Table3" runat="Server" width="100%">
                <asp:TableRow>
                    <asp:TableCell Wrap="False">
                        <asp:Label ID="Label1" runat="server" Text="Item(s) Booked" forecolor="Gray" font-size="X-Small" font-names="Arial" font-bold="True" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:table>
            <asp:DataGrid id="grd_BookingItems" runat="server" Width="90%" Font-Names="Verdana" Font-Size="XX-Small" GridLines="None" AutoGenerateColumns="False" CellSpacing="-1">
                <HeaderStyle font-bold="True"></HeaderStyle>
                <AlternatingItemStyle forecolor="WhiteSmoke"></AlternatingItemStyle>
                <ItemStyle forecolor="#0000C0"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="ProductCode" HeaderText="Code">
                        <HeaderStyle wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDate" HeaderText="Prod Date">
                        <HeaderStyle wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" HeaderText="Description">
                        <HeaderStyle wrap="False" forecolor="Gray"></HeaderStyle>
                        <ItemStyle forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ItemsOut" HeaderText="Qty" DataFormatString="{0:#,##0}">
                        <HeaderStyle wrap="False" horizontalalign="Right" forecolor="Gray"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Right" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="UnitValue" HeaderText="Unit Cost" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle wrap="False" horizontalalign="Right" forecolor="Gray"></HeaderStyle>
                        <ItemStyle horizontalalign="Right" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Cost" HeaderText="Total Cost" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle wrap="False" horizontalalign="Right" forecolor="Gray"></HeaderStyle>
                        <ItemStyle horizontalalign="Right" forecolor="Navy" verticalalign="Top"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
          </asp:Panel>
        </asp:Panel>
        <asp:Label id="lblError" runat="server" font-names="Verdana" font-size="X-Small" forecolor="#00C000"></asp:Label>
    </form>
    <script language="JavaScript" type="text/javascript" src="wz_tooltip.js"></script>
    <script language="JavaScript" type="text/javascript" src="library_functions.js"></script>
</body>
</html>
