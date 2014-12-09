<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    '   Cancelled Booking Report

    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Dim iDefaultHistoryPeriod As Integer = -3    'last 3 months

    Protected Sub Page_load(ByVal Source As Object, ByVal E As EventArgs)
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

            ShowMainPage()
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
    
    Protected Sub btnReselectFilterSettings_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectFilterSettings()
    End Sub
    
    Protected Sub ReselectFilterSettings()
        btnRunReport1.Visible = True
        btnRunReport2.Visible = True
        btnReselectFilterSettings1.Visible = False
        btnReselectFilterSettings2.Visible = False
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
        Call ShowMainPage()
    End Sub

    Protected Sub ShowMainPage()
        pnlAWBList.Visible = False
        pnlBookingDetail.Visible = False
    End Sub

    Protected Sub ShowAWBList()
        pnlAWBList.Visible = True
        pnlBookingDetail.Visible = False
    End Sub

    Protected Sub HideAWBList()
        pnlAWBList.Visible = False
        pnlBookingDetail.Visible = False
    End Sub

    Protected Sub ShowBookingDetail()
        pnlAWBList.Visible = False
        pnlBookingDetail.Visible = True
    End Sub

    Protected Sub GetDateRange()
        btnRunReport1.Visible = False
        btnRunReport2.Visible = False
        btnReselectFilterSettings1.Visible = True
        btnReselectFilterSettings2.Visible = True
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
            
            BindAWBGrid("BookedOn DESC")
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
    
    Protected Sub grd_AWB_item_click(ByVal s As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "info" Then
            Dim cell_BookingKey As TableCell = e.Item.Cells(1)
            Dim cell_AWB As TableCell = e.Item.Cells(2)
            If IsNumeric(cell_BookingKey.Text) Then
                lBookingKey = CLng(cell_BookingKey.Text)
            Else
                lBookingKey = 0
            End If

            If lBookingKey > 0 Then
                ResetBookingDetailForm()
                GetBookingDetailFromKey(lBookingKey)
                BindStockItems("ProductCode")
                ShowBookingDetail()
            End If
        End If
    End Sub

    Protected Sub btn_GetCancelledBookings_click(ByVal s As Object, ByVal e As EventArgs)
        BindAWBGrid("BookedOn DESC")
        ShowAWBList()
    End Sub

    Protected Sub btn_BackToList_click(ByVal s As Object, ByVal e As EventArgs)
        ShowAWBList()
    End Sub

    Protected Sub BindAWBGrid(ByVal SortField As String)
        lblError.Text = ""
        lbl_AWBList.Text = ""
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter

        oAdapter.SelectCommand = New SqlCommand("spASPNET_Report_CancelledBookings2", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = CDate(sFromDate)

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = CDate(sToDate)

        Try
            oAdapter.Fill(oDataSet, "AWBs")
            Dim Source As DataView = oDataSet.Tables("AWBs").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                grd_AWBs.DataSource = Source
                grd_AWBs.DataBind()
                grd_AWBs.Visible = True
            Else
                grd_AWBs.Visible = False
                lbl_AWBList.Text = "no records found"
                lblReportGeneratedDateTime.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub BindStockItems(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_LogisticBooking_GetMovements", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@StockBookingKey", SqlDbType.Int, 4))
        oAdapter.SelectCommand.Parameters("@StockBookingKey").Value = lBookingKey  'from viewstate
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
                lblError.Text = "... no data found for this date range"
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

    Protected Sub grd_AWB_Page_Change(ByVal s As Object, ByVal e As DataGridPageChangedEventArgs)
        grd_AWBs.CurrentPageIndex = e.NewPageIndex
        BindAWBGrid("BookedOn DESC")
    End Sub

    Protected Sub GetBookingDetailFromKey(ByVal lBookingKey As Long)
        Dim sCnorAddr1 As String = String.Empty
        Dim sCnorAddr2 As String = String.Empty
        Dim sCnorAddr3 As String = String.Empty
        Dim sCnorAddr4 As String = String.Empty
        Dim sCnorAddr5 As String = String.Empty
        Dim sCneeAddr1 As String = String.Empty
        Dim sCneeAddr2 As String = String.Empty
        Dim sCneeAddr3 As String = String.Empty
        Dim sCneeAddr4 As String = String.Empty
        Dim sCneeAddr5 As String = String.Empty
        Dim sPOD As String = String.Empty
        Dim sAgentDetails As String = String.Empty

        Dim oDataReader As SqlDataReader
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_LogisticBooking_GetDetailsFromKey2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As New SqlParameter("@StockBookingKey", SqlDbType.Int, 4)
        oCmd.Parameters.Add(oParam)
        oParam.Value = lBookingKey  'from viewstate
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If Not IsDBNull(oDataReader("CnorName")) Then
                sCnorAddr1 = oDataReader("CnorName")
            End If
            If Not IsDBNull(oDataReader("CnorAddr1")) Then
                sCnorAddr2 = oDataReader("CnorAddr1")
            End If
            If Not IsDBNull(oDataReader("CnorAddr2")) Then
                sCnorAddr3 = oDataReader("CnorAddr2") & " "
            End If
            If Not IsDBNull(oDataReader("CnorAddr3")) Then
                sCnorAddr3 &= oDataReader("CnorAddr3") & " "
            End If
            If Not IsDBNull(oDataReader("CnorTown")) Then
                sCnorAddr3 &= oDataReader("CnorTown")
            End If
            If Not IsDBNull(oDataReader("CnorState")) Then
                sCnorAddr4 = oDataReader("CnorState") & " "
            End If
            If Not IsDBNull(oDataReader("CnorPostCode")) Then
                sCnorAddr4 &= oDataReader("CnorPostCode") & " "
            End If
            If Not IsDBNull(oDataReader("CnorCountry")) Then
                sCnorAddr4 &= oDataReader("CnorCountry")
            End If
            If Not IsDBNull(oDataReader("CnorCtcName")) Then
                sCnorAddr5 = oDataReader("CnorCtcName")
            End If

            lblCnorAddr1.Text = sCnorAddr1
            lblCnorAddr2.Text = sCnorAddr2
            lblCnorAddr3.Text = sCnorAddr3
            lblCnorAddr4.Text = sCnorAddr4
            lblCnorAddr5.Text = sCnorAddr5

            If Not IsDBNull(oDataReader("CneeName")) Then
                sCneeAddr1 = oDataReader("CneeName")
            End If
            If Not IsDBNull(oDataReader("CneeAddr1")) Then
                sCneeAddr2 = oDataReader("CneeAddr1")
            End If
            If Not IsDBNull(oDataReader("CneeAddr2")) Then
                sCneeAddr3 = oDataReader("CneeAddr2") & " "
            End If
            If Not IsDBNull(oDataReader("CneeAddr3")) Then
                sCneeAddr3 &= oDataReader("CneeAddr3") & " "
            End If
            If Not IsDBNull(oDataReader("CneeTown")) Then
                sCneeAddr3 &= oDataReader("CneeTown")
            End If
            If Not IsDBNull(oDataReader("CneeState")) Then
                sCneeAddr4 = oDataReader("CneeState") & " "
            End If
            If Not IsDBNull(oDataReader("CneePostCode")) Then
                sCneeAddr4 &= oDataReader("CneePostCode") & " "
            End If
            If Not IsDBNull(oDataReader("CneeCountry")) Then
                sCneeAddr4 &= oDataReader("CneeCountry")
            End If
            If Not IsDBNull(oDataReader("CneeCtcName")) Then
                sCneeAddr5 = oDataReader("CneeCtcName")
            End If

            lblCneeAddr1.Text = sCneeAddr1
            lblCneeAddr2.Text = sCneeAddr2
            lblCneeAddr3.Text = sCneeAddr3
            lblCneeAddr4.Text = sCneeAddr4
            lblCneeAddr5.Text = sCneeAddr5

            If Not IsDBNull(oDataReader("POD")) Then
                lblPOD.Text = oDataReader("POD")
            End If

            If Not IsDBNull(oDataReader("BookedBy")) Then
                lblBookedBy.Text = oDataReader("BookedBy")
            End If

            If Not IsDBNull(oDataReader("BookedOn")) Then
                lblBookedOn.Text = Format(oDataReader("BookedOn"), "dd-MM-yy HH:mm")
            End If

            If Not IsDBNull(oDataReader("AWB")) Then
                lblAWB.Text = oDataReader("AWB")
            End If

            If Not IsDBNull(oDataReader("StateId")) Then
                lblStateId.Text = oDataReader("StateId")
            End If

            If Not IsDBNull(oDataReader("CustomerBookingReference")) Then
                lblBookingRef1.Text = oDataReader("CustomerBookingReference")
            End If

            If Not IsDBNull(oDataReader("CustomerBookingReference2")) Then
                lblBookingRef2.Text = oDataReader("CustomerBookingReference2")
            End If

            If Not IsDBNull(oDataReader("CustomerDepartmentId")) Then
                lblDepartmentId.Text = oDataReader("CustomerDepartmentId")
            End If

            If Not IsDBNull(oDataReader("SpecialInstructions")) Then
                lblSpecialInstructions.Text = oDataReader("SpecialInstructions")
            End If

            If Not IsDBNull(oDataReader("ShippingInformation")) Then
                lblShippingInformation.Text = oDataReader("ShippingInformation")
            End If

            If Not IsDBNull(oDataReader("TypeId")) Then
                lblTypeId.Text = oDataReader("TypeId")
            End If

            If Not IsDBNull(oDataReader("CashOnDelAmount")) Then
                lblAWBCost.Text = oDataReader("CashOnDelAmount")
            End If

            If Not IsDBNull(oDataReader("Weight")) Then
                lblWeight.Text = oDataReader("Weight")
            End If

            If Not IsDBNull(oDataReader("AgentName")) Then
                lblAgent.Text = oDataReader("AgentName")
            End If

        Catch ex As SqlException
            lblError.Text = ex.ToString
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

    Protected Sub ResetBookingDetailForm()
        lblCnorAddr1.Text = ""
        lblCnorAddr2.Text = ""
        lblCnorAddr3.Text = ""
        lblCnorAddr4.Text = ""
        lblCnorAddr5.Text = ""
        lblCneeAddr1.Text = ""
        lblCneeAddr2.Text = ""
        lblCneeAddr3.Text = ""
        lblCneeAddr4.Text = ""
        lblCneeAddr5.Text = ""
        lblPOD.Text = ""
        lblBookedBy.Text = ""
        lblBookedOn.Text = ""
        lblAWB.Text = ""
        lblStateId.Text = ""
        lblBookingRef1.Text = ""
        lblBookingRef2.Text = ""
        lblDepartmentId.Text = ""
        lblSpecialInstructions.Text = ""
        lblShippingInformation.Text = ""
        lblTypeId.Text = ""
        lblAWBCost.Text = ""
        lblWeight.Text = ""
        lblAgent.Text = ""
        lblStatus.Text = ""
    End Sub

    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = -1 Then
            ddlProductGroup.Items.RemoveAt(0)
        End If
        pnSelectedProductGroup = ddl.SelectedValue
        btnRunReport1.Enabled = True
        btnRunReport2.Enabled = True
        Call ReselectFilterSettings()
    End Sub
    
    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
        Call ReselectFilterSettings()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
    End Sub
    
    Property sToDate() As String
        Get
            Dim o As Object = ViewState("CBR_ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CBR_ToDate") = Value
        End Set
    End Property
    
    Property sFromDate() As String
        Get
            Dim o As Object = ViewState("CBR_FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("CBR_FromDate") = Value
        End Set
    End Property
    
    Property lBookingKey() As Long
        Get
            Dim o As Object = ViewState("CBR_BookingKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("CBR_BookingKey") = Value
        End Set
    End Property

    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("CBR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("CBR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("CBR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("CBR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("CBR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("CBR_IsProductOwner") = Value
        End Set
    End Property
   
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Cancelled Bookings Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="Form1" runat="server">
        <table id="tblDateRangeSelector" runat="server" visible="true">
            <tr id="Tr1" runat="server">
                <td colspan="4" style="white-space:nowrap">
                    <asp:Label ID="lblPageHeading" runat="server" ForeColor="silver" Font-Size="Small"
                        Font-Bold="True" Font-Names="Arial">
                             Cancelled Bookings Report</asp:Label><br />
                    <br />
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
                    <span class="informational dark">From:</span>
                        <asp:TextBox ID="tbFromDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1" src="../images/SmallCalendar.gif" runat="server" border="0" IE:visible="true" visible="false" alt="" /></a><span id="spnDateExample1" runat="server" visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2007)</span>
                </td>
                <td style="width: 265px; white-space:nowrap">
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                           <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img2" src="../images/SmallCalendar.gif" runat="server" border="0" IE:visible="true" visible="false" alt=""
                               /></a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2008)</span>
                </td>
                <td style="width: 253px">
                    <asp:Button ID="btnRunReport1" runat="server" Text="generate report" Visible="true" OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnReselectFilterSettings1" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectFilterSettings_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between calendar interface and dropdown interface"></asp:LinkButton></td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 265px; height: 26px;">
                    <span class="informational dark">From:</span> &nbsp;<asp:DropDownList ID="ddlFromDay"
                        runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                <td style="width: 265px; height: 26px;">
                    <span class="informational dark">To:</span> &nbsp;<asp:DropDownList ID="ddlToDay"
                        runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                <td style="width: 253px; height: 26px;">
                    <asp:Button ID="btnRunReport2" runat="server" Text="generate report" OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnReselectFilterSettings2" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectFilterSettings_Click" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px; height: 26px;">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface"></asp:LinkButton></td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate"
                        ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1"
                        ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate"
                        ErrorMessage=" - invalid format for expiry date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True"
                            ErrorMessage=" - expiry year before 2000, after 2020, or not a valid date!" Font-Names="Verdana"
                            Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface"
                            EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlAWBList" runat="server" Width="100%">
            <asp:DataGrid ID="grd_AWBs" runat="server" Font-Size="XX-Small" Font-Names="Arial"
                Width="100%" OnSortCommand="SortAWBListColumns" AllowSorting="True" GridLines="None"
                AutoGenerateColumns="False" OnPageIndexChanged="grd_AWB_Page_Change" PageSize="12"
                AllowPaging="False" OnItemCommand="grd_AWB_item_click">
                <HeaderStyle Font-Size="XX-Small" Font-Names="Arial" Wrap="False" BorderColor="Gray">
                </HeaderStyle>
                <PagerStyle NextPageText="" Font-Size="X-Small" Font-Names="Verdana" Font-Bold="True"
                    PrevPageText="" HorizontalAlign="Center" ForeColor="Blue" PageButtonCount="15"
                    Wrap="False" Mode="NumericPages"></PagerStyle>
                <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                <ItemStyle Font-Size="XX-Small" Font-Names="Arial" BackColor="LightGray"></ItemStyle>
                <Columns>
                    <asp:TemplateColumn>
                        <HeaderStyle ForeColor="Blue" Width="5%"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                        <HeaderTemplate>
                            Info
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:ImageButton runat="server" CommandName="info" ImageUrl="../images/icon_info.gif"
                                ToolTip="shipment details"></asp:ImageButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="LogisticBookingKey" SortExpression="Booking No" HeaderText="Booking">
                        <HeaderStyle Wrap="False" ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedOn" SortExpression="BookedOn" HeaderText="Date"
                        DataFormatString="{0:dd-MMM-yy}">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeName" SortExpression="CneeName" HeaderText="Company">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="City">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeCountryName" SortExpression="CneeCountryName" HeaderText="Country">
                        <HeaderStyle ForeColor="Blue"></HeaderStyle>
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <asp:Label ID="lbl_AWBList" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                ForeColor="Blue"></asp:Label>
            <br />
            <asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" Font-Size="XX-Small"
                Font-Names="Verdana, Sans-Serif" ForeColor="Green" Visible="false"></asp:Label>
        </asp:Panel>
        <asp:Panel ID="pnlBookingDetail" runat="server">
            <asp:Table ID="tblBookingDetail" runat="Server" Width="100%">
                <asp:TableRow>
                    <asp:TableCell Wrap="False">
                        <asp:Label runat="server" forecolor="silver" font-size="X-Small" font-names="Arial" font-bold="True">Booking
                        Details</asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Table ID="Table4" runat="server" Width="750px" Font-Size="XX-Small" Font-Names="Arial">
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Width="15%" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">From:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Width="35%" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCnorAddr1"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Width="15%" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">To:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Width="35%" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCneeAddr1"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCnorAddr2"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCneeAddr2"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCnorAddr3"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCneeAddr3"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCnorAddr4"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCneeAddr4"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCnorAddr5"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblCneeAddr5"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <asp:Table ID="Table5" runat="server" Font-Size="XX-Small" Font-Names="Arial" Width="750px">
                <asp:TableRow>
                    <asp:TableCell Width="15%" Wrap="False"></asp:TableCell>
                    <asp:TableCell Width="35%" Wrap="False"></asp:TableCell>
                    <asp:TableCell BackColor="PaleTurquoise" Width="15%" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Received by:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="PaleTurquoise" Width="35%" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblPOD"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <br />
            <asp:DataGrid ID="grd_BookingItems" runat="server" Width="750px" Font-Size="X-Small"
                Font-Names="Arial" GridLines="None" AutoGenerateColumns="False" CellSpacing="-1">
                <HeaderStyle Font-Bold="True"></HeaderStyle>
                <AlternatingItemStyle ForeColor="#0000C0"></AlternatingItemStyle>
                <ItemStyle ForeColor="#0000C0"></ItemStyle>
                <Columns>
                    <asp:BoundColumn DataField="ProductCode" HeaderText="Product Code">
                        <HeaderStyle Wrap="False" ForeColor="silver"></HeaderStyle>
                        <ItemStyle Wrap="False" ForeColor="#0000C0"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDate" HeaderText="Product Date">
                        <HeaderStyle Wrap="False" ForeColor="silver"></HeaderStyle>
                        <ItemStyle Wrap="False" ForeColor="#0000C0"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" HeaderText="Description">
                        <HeaderStyle ForeColor="silver"></HeaderStyle>
                        <ItemStyle ForeColor="#0000C0"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ItemsOut" HeaderText="Quantity" DataFormatString="{0:#,##0}">
                        <HeaderStyle Wrap="False" HorizontalAlign="Right" ForeColor="silver"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Right" ForeColor="#0000C0"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <br />
            <asp:Table ID="Table6" runat="server" Width="750px" Font-Size="XX-Small" Font-Names="Arial">
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Width="15%" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Booked by:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Width="35%" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblBookedBy"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender" Width="15%">
                        <asp:Label runat="server" forecolor="#0000C0">Booked on:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Width="35%">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblBookedOn"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Warehouse Status:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblStateId"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">Air Waybill:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblAWB"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Spcl Instr:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" ColumnSpan="3" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblSpecialInstructions"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Packing Note:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" ColumnSpan="3" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblShippingInformation"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Agent:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblAgent"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">NOP / Weight:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblWeight"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Booking Ref:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro" Wrap="False">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblBookingRef1"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">PCID:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblDepartmentId"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Rating:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblBookingRef2"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender">
                        <asp:Label runat="server" forecolor="#0000C0">RO:</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblTypeId"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
                <asp:TableRow>
                    <asp:TableCell BackColor="Lavender" Wrap="False">
                        <asp:Label runat="server" forecolor="#0000C0">Shipping cost (£):</asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro">
                        <asp:Label runat="server" ForeColor="#0000C0" ID="lblAWBCost"></asp:Label>
                    </asp:TableCell>
                    <asp:TableCell BackColor="Lavender"></asp:TableCell>
                    <asp:TableCell BackColor="Gainsboro"></asp:TableCell>
                </asp:TableRow>
                <asp:TableRow Height="15px">
                    <asp:TableCell ColumnSpan="4" HorizontalAlign="Right">
                        <asp:LinkButton runat="server" ForeColor="Blue" CausesValidation="False" Font-Size="X-Small"
                            OnClick="btn_BackToList_click">back&nbsp;to&nbsp;list</asp:LinkButton>
                        <br />
                        <asp:Label ID="lblStatus" runat="server" Font-Size="X-Small" Font-Names="Arial" ForeColor="#00C000"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
        </asp:Panel>
        <br />
        <asp:Label ID="lblError" runat="server" ForeColor="red" Font-Names="Arial" Font-Size="XX-Small"></asp:Label>
    </form>
</body>
</html>
