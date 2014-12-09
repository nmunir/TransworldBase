<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    '   Consignment Cost Report
    
    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Dim iDefaultHistoryPeriod As Integer = -1    'last 1 months
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If

        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        If Not IsPostBack Then
            pbIsProductOwner = CBool(Session("UserType").ToString.ToLower.Contains("owner"))
            Call GetSiteFeatures()
            trProductGroups.Visible = pbProductOwners
            ' pbProductOwners = site-wide Product Owners attribute; pbIsProductOwner = this user
            If pbIsProductOwner Then
                If pbProductOwners Then
                    ddlProductGroup.Visible = True
                    Call PopulateProductGroups(Session("UserKey"))
                    btnShowProductGroups.Visible = False
                Else
                    WebMsgBox.Show("Cannot show report as Product Owners attribute is not enabled for this web site")
                    Exit Sub
                End If
            Else
                If pbProductOwners Then
                    btnShowProductGroups.Visible = True
                Else
                    btnShowProductGroups.Visible = False
                End If
                pnSelectedProductGroup = 0
            End If
            lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
            Dim dteLastYear As Date = Date.Today.AddMonths(iDefaultHistoryPeriod)
            tbFromDate.Text = dteLastYear.ToString("dd-MMM-yy")
            tbToDate.Text = Now.ToString("dd-MMM-yy")
            
            Dim sYear As String = Year(Now)
            Dim i As Integer
            For i = CInt(sYear) To CInt(sYear) - 6 Step -1
                ddlToYear.Items.Add(i.ToString)
                ddlFromYear.Items.Add(i.ToString)
            Next
            ddlToYear.Items(0).Selected = True
            ddlFromYear.Items(1).Selected = True
            
            ddlToMonth.Items(Month(Now) - 1).Selected = True
            ddlFromMonth.Items(Month(Now) - 1).Selected = True

            ddlToDay.Items(Day(Now) - 1).Selected = True
            ddlFromDay.Items(Day(Now) - 1).Selected = True

            lnkbtnToggleSelectionStyle1.Text = STYLENAME_DROPDOWN
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_DROPDOWN
        End If
    End Sub
    
    Protected Sub GetSiteFeatures()
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
        pbProductOwners = dr("ProductOwners")
    End Sub
    
    Protected Sub PopulateProductGroups(ByVal nProductOwner As Integer)
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_Product_GetGroupsForOwner", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim paramProductOwner As SqlParameter = New SqlParameter("@ProductOwner", SqlDbType.Int)
        paramProductOwner.Value = nProductOwner
        oCmd.Parameters.Add(paramProductOwner)
       
        Dim paramCustomerKey As SqlParameter = New SqlParameter("@CustomerKey", SqlDbType.Int)
        paramCustomerKey.Value = Session("CustomerKey")
        oCmd.Parameters.Add(paramCustomerKey)
       
        Try
            oConn.Open()
            Dim oSqlDataReader As SqlDataReader = oCmd.ExecuteReader
            If oSqlDataReader.HasRows Then
                ddlProductGroup.Items.Add(New ListItem("- select product group -", -1))
                If Not pbIsProductOwner Then
                    ddlProductGroup.Items.Add(New ListItem("- all products -", 0))
                End If
                While oSqlDataReader.Read()
                    ddlProductGroup.Items.Add(New ListItem(oSqlDataReader("ProductGroupName"), oSqlDataReader("ProductGroupKey")))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("PopulateProductgGroupsDropdown: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        
        If ddlProductGroup.Items.Count <= 2 Then
            lblProductGroup.Text = "Product group: " & ddlProductGroup.Items(1).Text
            pnSelectedProductGroup = ddlProductGroup.Items(1).Value
            ddlProductGroup.Visible = False
        Else
            btnRunReport1.Enabled = False
            btnRunReport2.Enabled = False
            pnSelectedProductGroup = -1
        End If
    End Sub
    
    Protected Sub btnReselectReportFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectReportFilter()
    End Sub
    
    Protected Sub ReselectReportFilter()
        pnlAWBList.Visible = False
        btnRunReport1.Visible = True
        btnRunReport2.Visible = True
        btnReselectReportFilter1.Visible = False
        btnReselectReportFilter2.Visible = False
        tbFromDate.Enabled = True
        tbToDate.Enabled = True
        ddlFromDay.Enabled = True
        ddlFromMonth.Enabled = True
        ddlFromYear.Enabled = True
        ddlToDay.Enabled = True
        ddlToMonth.Enabled = True
        ddlToYear.Enabled = True
        spnDateExample1.Visible = True
        spnDateExample2.Visible = True
        imgCalendarButton1.Visible = True
        imgCalendarButton2.Visible = True
        ddlProductGroup.Enabled = True
    End Sub

    Protected Sub ShowAWBList()
        tblDateRangeSelector.Visible = True
        pnlAWBList.Visible = True
        pnlAWBDetail.Visible = False
    End Sub
    
    Protected Sub ShowAWBDetail()
        tblDateRangeSelector.Visible = False
        pnlAWBList.Visible = False
        pnlAWBDetail.Visible = True
    End Sub

    Protected Sub GetDateRange()
        btnRunReport1.Visible = False
        btnRunReport2.Visible = False
        btnReselectReportFilter1.Visible = True
        btnReselectReportFilter2.Visible = True
        tbFromDate.Enabled = False
        tbToDate.Enabled = False
        ddlFromDay.Enabled = False
        ddlFromMonth.Enabled = False
        ddlFromYear.Enabled = False
        ddlToDay.Enabled = False
        ddlToMonth.Enabled = False
        ddlToYear.Enabled = False
        
        If CalendarInterface.Visible Then
            sToDate = tbToDate.Text
            sFromDate = tbFromDate.Text
        Else
            sFromDate = ddlFromDay.SelectedItem.Text & "-" & ddlFromMonth.SelectedItem.Text & "-" & ddlFromYear.SelectedItem.Text
            sToDate = ddlToDay.SelectedItem.Text & "-" & ddlToMonth.SelectedItem.Text & "-" & ddlToYear.SelectedItem.Text
            tbFromDate.Text = sFromDate
            tbToDate.Text = sToDate
        End If
    End Sub
    
    Protected Sub btnRunReport_Click(ByVal s As Object, ByVal e As EventArgs)
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        
        If CalendarInterface.Visible Then
            Page.Validate("CalendarInterface")
        Else
            Dim sDate = ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text
            If IsDate(sDate) Then
            End If
        End If

        If (CalendarInterface.Visible And Page.IsValid) _
         Or (DropdownInterface.Visible _
          And IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) _
          And IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text)) Then
            Call GetDateRange()
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False

            lblReportGeneratedDateTime.Visible = True
            BindAWBGrid("AWB")
            ShowAWBList()
        Else
            If DropdownInterface.Visible Then
                If Not IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) Then
                    lblFromErrorMessage.Text = "Invalid date"
                End If
                If Not IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text) Then
                    lblToErrorMessage.Text = "Invalid date"
                End If
            End If
        End If
        ddlProductGroup.Enabled = False
    End Sub
    
    Protected Sub btn_BackToList_click(ByVal s As Object, ByVal e As EventArgs)
        ShowAWBList()
    End Sub
    
    Protected Sub grd_AWB_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        Dim cell_AWB As TableCell = e.Item.Cells(1)
        sAWBNumber = CStr(cell_AWB.Text)
        If sAWBNumber <> "" Then
            ResetAWBDetailForm()
            GetAWBDetailFromAWBNo(sAWBNumber)
            ShowAWBDetail()
        End If
    End Sub
    
    Protected Sub BindAWBGrid(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Report_AWBCosts3", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = CDate(sFromDate)

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = DateAdd("D", 1, CDate(sToDate))

        Try
            oAdapter.Fill(oDataSet, "Movements")
            Dim Source As DataView = oDataSet.Tables("Movements").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                grd_AWBs.Visible = True
                grd_AWBs.DataSource = Source
                grd_AWBs.DataBind()
            Else
                grd_AWBs.Visible = False
                lblError.Text = "... no data found for these filter settings"
                lblReportGeneratedDateTime.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SortAWBListColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        BindAWBGrid(E.SortExpression)
    End Sub
    
    Protected Sub GetAWBDetailFromAWBNo(ByVal sAWBNumber As String)
        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Consignment_GetDetailsFromAWB2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam1 As New SqlParameter("@AWB", SqlDbType.NVarChar, 50)
        oCmd.Parameters.Add(oParam1)
        oParam1.Value = sAWBNumber
        Dim oParam2 As New SqlParameter("@CustomerKey", SqlDbType.Int)
        oCmd.Parameters.Add(oParam2)
        oParam2.Value = Session("CustomerKey")
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
    
            If Not IsDBNull(oDataReader("AWB")) Then
                lbl_AWB.Text = oDataReader("AWB")
            Else
                lbl_AWB.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("ShipDate")) Then
                lbl_Date.Text = Format(oDataReader("ShipDate"), "d-MMM-yyyy")
            Else
                lbl_Date.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CneeName")) Then
                lbl_CneeName.Text = oDataReader("CneeName")
            Else
                lbl_CneeName.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CneeAddr1")) Then
                lbl_CneeAddr1.Text = oDataReader("CneeAddr1")
            Else
                lbl_CneeAddr1.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CneeAddr2")) Then
                lbl_CneeAddr2.Text = oDataReader("CneeAddr2")
            Else
                lbl_CneeAddr2.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CneeAddr3")) Then
                lbl_CneeAddr3.Text = oDataReader("CneeAddr3")
            Else
                lbl_CneeAddr3.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CneeAddr4")) Then
                lbl_CneeAddr4.Text = oDataReader("CneeAddr4")
            Else
                lbl_CneeAddr4.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CustomerRef1")) Then
                lbl_Ref1.Text = oDataReader("CustomerRef1")
            Else
                lbl_Ref1.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("Weight")) Then
                lbl_Weight.Text = oDataReader("Weight")
            Else
                lbl_Weight.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("Description")) Then
                lbl_Description.Text = oDataReader("Description")
            Else
                lbl_Description.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("POD")) Then
                lbl_POD.Text = oDataReader("POD")
            Else
                lbl_POD.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("Misc1")) Then
                lbl_Misc1.Text = oDataReader("Misc1")
            Else
                lbl_Misc1.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("Misc2")) Then
                lbl_Misc2.Text = oDataReader("Misc2")
            Else
                lbl_Misc2.Text = " "
            End If
    
            If Not IsDBNull(oDataReader("CashOnDelAmount")) Then
                lbl_Cost.Text = "£ " & Format(oDataReader("CashOnDelAmount"), "#,###.#0")
            Else
                lbl_Cost.Text = " "
            End If
    
        Catch ex As SqlException
            lblAWBDetailError.Text = ex.ToString
        Finally
            oDataReader.Close()
            oConn.Close()
        End Try
    
    End Sub
    
    Protected Sub lnkbtnToggleSelectionStyle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CalendarInterface.Visible = True Then
            CalendarInterface.Visible = False
            DropdownInterface.Visible = True
            If Page.IsValid Then
                Dim dDate As Date
                Dim nVal As Integer
                If IsDate(tbFromDate.Text) Then
                    dDate = Date.Parse(tbFromDate.Text)
                    nVal = dDate.Day
                    ddlFromDay.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Month)
                    ddlFromMonth.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Year)
                    For i As Integer = 0 To ddlFromYear.Items.Count - 1
                        If ddlFromYear.Items(i).Text = CStr(nVal) Then
                            ddlFromYear.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If

                If IsDate(tbToDate.Text) Then
                    dDate = Date.Parse(tbToDate.Text)
                    nVal = dDate.Day
                    ddlToDay.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Month)
                    ddlToMonth.SelectedIndex = nVal - 1
                    nVal = CStr(dDate.Year)
                    For i As Integer = 0 To ddlToYear.Items.Count - 1
                        If ddlToYear.Items(i).Text = CStr(nVal) Then
                            ddlToYear.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If
            End If
        Else
            CalendarInterface.Visible = True
            DropdownInterface.Visible = False
            Dim arrMonths() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
            If IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) Then
                tbFromDate.Text = ddlFromDay.SelectedValue & "-" & arrMonths(ddlFromMonth.SelectedIndex) & "-" & ddlFromYear.SelectedValue
            End If
            If IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text) Then
                tbToDate.Text = ddlToDay.SelectedValue & "-" & arrMonths(ddlToMonth.SelectedIndex) & "-" & ddlToYear.SelectedValue
            End If
        End If
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        If lnkbtnToggleSelectionStyle1.Text = STYLENAME_CALENDAR Then
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_DROPDOWN
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_DROPDOWN
        Else
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_CALENDAR
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_CALENDAR
        End If
    End Sub

    Protected Sub ResetAWBDetailForm()
        lbl_AWB.Text = ""
        lbl_Date.Text = ""
        lbl_CneeName.Text = ""
        lbl_CneeAddr1.Text = ""
        lbl_CneeAddr2.Text = ""
        lbl_CneeAddr3.Text = ""
        lbl_CneeAddr4.Text = ""
        lbl_Ref1.Text = ""
        lbl_Service.Text = ""
        lbl_Weight.Text = ""
        lbl_Description.Text = ""
        lbl_POD.Text = ""
        lbl_Misc1.Text = ""
        lbl_Misc2.Text = ""
        lbl_Cost.Text = ""
    End Sub

    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = -1 Then
            ddlProductGroup.Items.RemoveAt(0)
        End If
        pnSelectedProductGroup = ddl.SelectedValue
        btnRunReport1.Enabled = True
        btnRunReport2.Enabled = True
        Call ReselectReportFilter()
    End Sub
    
    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
        Call ReselectReportFilter()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
    End Sub
    
    Property lBookingKey() As Long
        Get
            Dim o As Object = ViewState("BookingKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("BookingKey") = Value
        End Set
    End Property
    
    Property sAWBNumber() As String
        Get
            Dim o As Object = ViewState("AWBNumber")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("AWBNumber") = Value
        End Set
    End Property
    
    Property sToDate() As String
        Get
            Dim o As Object = ViewState("ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ToDate") = Value
        End Set
    End Property
    
    Property sFromDate() As String
        Get
            Dim o As Object = ViewState("FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("FromDate") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("CCR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CCR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("CCR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("CCR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("CCR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("CCR_IsProductOwner") = Value
        End Set
    End Property
   
    Protected Sub grd_AWBs_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Const CUSTREF_GRID_OFFSET As Integer = 6     ' 6 - 9
        Dim dgiea As DataGridItemEventArgs = e
        If Not cbIncludeCustRef1.Checked Then
            dgiea.Item.Cells(CUSTREF_GRID_OFFSET).Visible = False
        End If
        If Not cbIncludeCustRef2.Checked Then
            dgiea.Item.Cells(CUSTREF_GRID_OFFSET + 1).Visible = False
        End If
        If Not cbIncludeCustRef3.Checked Then
            dgiea.Item.Cells(CUSTREF_GRID_OFFSET + 2).Visible = False
        End If
        If Not cbIncludeCustRef4.Checked Then
            dgiea.Item.Cells(CUSTREF_GRID_OFFSET + 3).Visible = False
        End If
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Consignment Cost Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmConsignmentCostreport" runat="server">
        <table id="tblDateRangeSelector" runat="server" visible="true">
            <tr id="Tr1" runat="server">
                <td colspan="4" style="white-space:nowrap">
                  <asp:Label ID="lblPageHeading" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial">Consignment Cost Report</asp:Label>
                  <br /><br />
                </td>
            </tr>
            <tr runat="server" visible="true">
                <td colspan="2" style="white-space: nowrap">
                </td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
            <tr runat="server" visible="true" id="trProductGroups">
                <td colspan="2" style="white-space: nowrap">
                    &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                        Visible="False">
                    </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Names="Verdana"
                        Font-Size="X-Small" Font-Bold="True"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                    <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                        Text="show product groups" Visible="False" /></td>
            </tr>
            <tr runat="server" visible="true">
                <td style="width: 265px; white-space: nowrap">
                </td>
                <td style="width: 265px; white-space: nowrap">
                </td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 265px; white-space:nowrap">
                    <span class="informational dark">&nbsp;From:</span>
                        <asp:TextBox ID="tbFromDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1"
                                 src="../images/SmallCalendar.gif"
                                 runat="server"
                                 border="0"
                              IE:visible="true"
                                 visible="false" alt=""
                               /></a>
                    <span id="spnDateExample1" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2007)</span>
                    </td>
                <td style="width: 265px; white-space:nowrap">
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                           <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img2" src="../images/SmallCalendar.gif" runat="server" border="0" IE:visible="true" visible="false" alt=""
                               /></a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2008)</span>
                   </td>
                <td style="width: 253px">
                <asp:Button ID="btnRunReport1"
                     runat="server"
                     Text="generate report"
                     Visible="true"
                     OnClick="btnRunReport_Click" Width="170px" />
                <asp:Button ID="btnReselectReportFilter1"
                     runat="server"
                     Text="re-select report filter"
                     Visible="false"
                      OnClick="btnReselectReportFilter_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click" ToolTip="toggles between calendar interface and dropdown interface"></asp:LinkButton></td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 265px">
                    <span class="informational dark">&nbsp;From:</span>
                    &nbsp;<asp:DropDownList ID="ddlFromDay" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>01</asp:ListItem><asp:ListItem>02</asp:ListItem>
                        <asp:ListItem>03</asp:ListItem>
                        <asp:ListItem>04</asp:ListItem>
                        <asp:ListItem>05</asp:ListItem>
                        <asp:ListItem>06</asp:ListItem>
                        <asp:ListItem>07</asp:ListItem>
                        <asp:ListItem>08</asp:ListItem>
                        <asp:ListItem>09</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>13</asp:ListItem>
                        <asp:ListItem>14</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>17</asp:ListItem>
                        <asp:ListItem>18</asp:ListItem>
                        <asp:ListItem>19</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>21</asp:ListItem>
                        <asp:ListItem>22</asp:ListItem>
                        <asp:ListItem>23</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>26</asp:ListItem>
                        <asp:ListItem>27</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>29</asp:ListItem>
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>31</asp:ListItem>
                        </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>Jan</asp:ListItem>
                        <asp:ListItem>Feb</asp:ListItem>
                        <asp:ListItem>Mar</asp:ListItem>
                        <asp:ListItem>Apr</asp:ListItem>
                        <asp:ListItem>May</asp:ListItem>
                        <asp:ListItem>Jun</asp:ListItem>
                        <asp:ListItem>Jul</asp:ListItem>
                        <asp:ListItem>Aug</asp:ListItem>
                        <asp:ListItem>Sep</asp:ListItem>
                        <asp:ListItem>Oct</asp:ListItem>
                        <asp:ListItem>Nov</asp:ListItem>
                        <asp:ListItem>Dec</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>&nbsp;</td>
                <td style="width: 265px">
                    <span class="informational dark">To:</span>
                    &nbsp;<asp:DropDownList ID="ddlToDay" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>01</asp:ListItem>
                        <asp:ListItem>02</asp:ListItem>
                        <asp:ListItem>03</asp:ListItem>
                        <asp:ListItem>04</asp:ListItem>
                        <asp:ListItem>05</asp:ListItem>
                        <asp:ListItem>06</asp:ListItem>
                        <asp:ListItem>07</asp:ListItem>
                        <asp:ListItem>08</asp:ListItem>
                        <asp:ListItem>09</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>13</asp:ListItem>
                        <asp:ListItem>14</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>17</asp:ListItem>
                        <asp:ListItem>18</asp:ListItem>
                        <asp:ListItem>19</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>21</asp:ListItem>
                        <asp:ListItem>22</asp:ListItem>
                        <asp:ListItem>23</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>26</asp:ListItem>
                        <asp:ListItem>27</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>29</asp:ListItem>
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>31</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                        <asp:ListItem>Jan</asp:ListItem>
                        <asp:ListItem>Feb</asp:ListItem>
                        <asp:ListItem>Mar</asp:ListItem>
                        <asp:ListItem>Apr</asp:ListItem>
                        <asp:ListItem>May</asp:ListItem>
                        <asp:ListItem>Jun</asp:ListItem>
                        <asp:ListItem>Jul</asp:ListItem>
                        <asp:ListItem>Aug</asp:ListItem>
                        <asp:ListItem>Sep</asp:ListItem>
                        <asp:ListItem>Oct</asp:ListItem>
                        <asp:ListItem>Nov</asp:ListItem>
                        <asp:ListItem>Dec</asp:ListItem>
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>&nbsp;</td>
                <td style="width: 253px">
                <asp:Button ID="btnRunReport2"
                     runat="server"
                     Text="generate report"
                      OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnReselectReportFilter2"
                     runat="server"
                     Text="re-select report filter"
                     Visible="false"
                      OnClick="btnReselectReportFilter_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                      </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                       ToolTip="toggles between calendar interface and dropdown interface"></asp:LinkButton></td>
            </tr>
            <tr>
                <td colspan="4">
                    &nbsp;Include Cust Ref:
                    <asp:CheckBox ID="cbIncludeCustRef1" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="1" />
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="cbIncludeCustRef2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="2" />
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="cbIncludeCustRef3" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="3" />
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="cbIncludeCustRef4" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="4" /></td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate" ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                            Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
        </table>
        <asp:Panel id="pnlAWBList" runat="server" Width="100%">
            <asp:DataGrid id="grd_AWBs" runat="server" Font-Size="XX-Small" Font-Names="Arial" Width="100%" OnSortCommand="SortAWBListColumns" AllowSorting="True" GridLines="None" AutoGenerateColumns="False" CellSpacing="-1" OnItemCommand="grd_AWB_item_click" OnItemDataBound="grd_AWBs_ItemDataBound">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" wrap="False" forecolor="Blue" bordercolor="Gray"></HeaderStyle>
                <PagerStyle nextpagetext="" font-size="X-Small" font-names="Verdana" font-bold="True" prevpagetext="" horizontalalign="Center" forecolor="Blue" pagebuttoncount="15" wrap="False" mode="NumericPages"></PagerStyle>
                <ItemStyle font-size="XX-Small" font-names="Arial"></ItemStyle>
                <Columns>
                    <asp:TemplateColumn HeaderText="Info">
                        <ItemTemplate>
                            <asp:ImageButton id="ImageButton1" runat="server" ImageUrl="../images/icon_info.gif"></asp:ImageButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="AWB" SortExpression="AWB" HeaderText="Air Waybill">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ShipDate" SortExpression="ShipDate" HeaderText="Date" DataFormatString="{0:dd-MMM-yy}">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeName" SortExpression="CneeName" HeaderText="Company">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="City">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CustomerRef1" SortExpression="CustomerRef1" HeaderText="Cust Ref 1">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CustomerRef2" SortExpression="CustomerRef2" HeaderText="Cust Ref 2">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Misc1" SortExpression="Misc1" HeaderText="Cust Ref 3">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Misc2" SortExpression="Misc2" HeaderText="Cust Ref 4">
                        <HeaderStyle forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CashOnDelAmount" SortExpression="CashOnDelAmount" HeaderText="Cost (&#163;)" DataFormatString="{0:#,##0.00}">
                        <HeaderStyle horizontalalign="Right" forecolor="Blue"></HeaderStyle>
                        <ItemStyle wrap="False" horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <br />
            &nbsp;<asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green" Visible="false"></asp:Label>
        </asp:Panel>
        <!-- Panel: AWB Detail  -->
        <asp:Panel id="pnlAWBDetail" runat="server" Visible="false" Width="100%">
            <asp:table id="Table10" runat="Server" width="600px">
                <asp:TableRow>
                    <asp:TableCell Wrap="False">
                        <asp:Label runat="server" forecolor="silver" font-size="X-Small" font-names="Arial" font-bold="True">Consignment
                        Detail</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell Width="1%">
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" VerticalAlign="Top">
                        <asp:LinkButton runat="server" CausesValidation="False" onclick="btn_BackToList_click">back to list</asp:LinkButton>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:table>
            <br />
            <asp:Table id="Table11" runat="server" Font-Size="XX-Small" Font-Names="Arial" Width="600px">
                <asp:TableRow Height="15px">
                    <asp:TableCell Width="5%"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Width="20%">
                        <asp:Label runat="server" forecolor="#0000C0">AWB:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Width="75%">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_AWB"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Date:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Date"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Consignee:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_CneeName"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_CneeAddr1"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_CneeAddr2"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_CneeAddr3"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_CneeAddr4"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Reference:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Ref1"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Misc 1:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Misc1"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Misc 2:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Misc2"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Service:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Service"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">NOP / Weight:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Weight"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Description:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Description"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Shipping Cost:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_Cost"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell BackColor="PaleTurquoise">
                        <asp:Label runat="server" forecolor="#0000C0">Received by:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="PaleTurquoise">
                        <asp:Label runat="server" forecolor="#0000C0" id="lbl_POD"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell></asp:TableCell>
                    <asp:TableCell HorizontalAlign="Right"></asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Label id="lblAWBDetailError" runat="server" forecolor="#0000C0"></asp:Label>
        </asp:Panel>
        <br />
        <asp:Label id="lblError" runat="server" forecolor="#00C000" font-names="Arial" font-size="XX-Small"></asp:Label>
    </form>
</body>
</html>
