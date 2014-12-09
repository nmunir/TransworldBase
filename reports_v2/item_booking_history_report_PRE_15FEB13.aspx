<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    ' Item Booking History Report
    
    ' USES "~/images/greendown.gif" & "~/images/orangeup.gif"
    ' STILL A WEIRD PROBLEM WITH REVERSE SORTING ON DATE FIELD, WHERE THE DATASET APPEARS TO BE OUT OF ORDER.

    ' ENHANCEMENTS OUTSTANDING AS AT 12JUN08
    
    ' allow selection of fields to be displayed
    ' sort excel data according to sort order of displayed data
    ' add custref3, custref4, other fields
    ' provide search box
    ' show statistics counts
    ' provide link to track & trace
    ' enlarge pager font
    ' fix non IE browser display problems
    
    ' LATER...
    ' modify stored procedure to return only records required

    Dim iDefaultHistoryPeriod As Integer = -3    'last 3 months
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"
    Const DEFAULT_SORT_ITEM As String = "Date"
    Const DEFAULT_SORT_DIRECTION As String = "ASC"
    
    Dim gbSortItemChanged As Boolean = False
    
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
                    PopulateProductGroups(Session("UserKey"))
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
            Dim dtFromDate As Date = Date.Today.AddMonths(iDefaultHistoryPeriod)
            Dim nVal As Integer
            tbFromDate.Text = dtFromDate.ToString("dd-MMM-yy")
            tbToDate.Text = Now.ToString("dd-MMM-yy")
            
            Dim sYear As String = Year(Now)
            Dim i As Integer
            For i = CInt(sYear) To CInt(sYear) - 6 Step -1
                ddlToYear.Items.Add(i.ToString)
                ddlFromYear.Items.Add(i.ToString)
            Next
            
            ddlToYear.Items(0).Selected = True
            ddlToMonth.Items(Month(Now) - 1).Selected = True
            ddlToDay.Items(Day(Now) - 1).Selected = True

            nVal = dtFromDate.Day
            ddlFromDay.SelectedIndex = nVal - 1
            nVal = CStr(dtFromDate.Month)
            ddlFromMonth.SelectedIndex = nVal - 1
            nVal = CStr(dtFromDate.Year)
            For i = 0 To ddlFromYear.Items.Count - 1
                If ddlFromYear.Items(i).Text = CStr(nVal) Then
                    ddlFromYear.SelectedIndex = i
                    Exit For
                End If
            Next
            lnkbtnToggleSelectionStyle1.Text = STYLENAME_DROPDOWN
            lnkbtnToggleSelectionStyle2.Text = STYLENAME_DROPDOWN
        End If
        If pbIsDisplayingData Then
            PopulateItemGridView()
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
    
    Protected Sub btnRunReport_Click(ByVal s As Object, ByVal e As EventArgs)
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        
        If cbFilterByProduct.Checked AndAlso ddlProductField.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select a field to filter by.")
            Exit Sub
        End If
        
        If CalendarInterface.Visible Then
            Page.Validate("CalendarInterface")
        Else
            Dim sDate = ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text
            If IsDate(sDate) Then
                
            End If
        End If

        If (CalendarInterface.Visible And Page.IsValid) Or (DropdownInterface.Visible _
          And IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) _
          And IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text)) Then
            Call GetDateRange()
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            lblReportGeneratedDateTime.Visible = True
            Call PopulateItemGridView()
            pnlData.Visible = True
            pbIsDisplayingData = True

            cbShowConsignmentCost.Enabled = False
            cbFilterByProduct.Enabled = False
            ddlProductField.Enabled = False
            ddlProductFieldValue.Enabled = False
                        
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
    
    Protected Sub PopulateItemGridView()
        Dim oDataTable As New DataTable
        oDataTable = dtGetData()
        gvItems.DataSource = oDataTable
        gvItems.PageSize = CInt(ddlRows.SelectedItem.Text)
        gvItems.DataBind()
        If oDataTable.Rows.Count > 0 Then
            btnExportToExcel1.Visible = True
            btnExportToExcel2.Visible = True
        Else
            btnExportToExcel1.Visible = False
            btnExportToExcel2.Visible = False
            lblError.Text = "No data found for this date range"
            lblReportGeneratedDateTime.Visible = False
        End If
    End Sub

    Protected Function dtGetData() As DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Report_ItemBookingHistory3", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@IncludeConsignmentCost", SqlDbType.Bit))
        oAdapter.SelectCommand.Parameters("@IncludeConsignmentCost").Value = cbShowConsignmentCost.Checked
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = CDate(psFromDate)
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = DateAdd("D", 1, CDate(psToDate))
    
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        
        If cbFilterByProduct.Checked Then
            Dim oFilteredDataTable As New DataTable
            oFilteredDataTable = oDataTable.Clone
            For Each dr As DataRow In oDataTable.Rows
                Dim drProduct As DataRow = ExecuteQueryToDataTable("SELECT ProductCategory FROM LogisticProduct WHERE LogisticProductKey = " & dr("LogisticProductKey")).Rows(0)
                If ddlProductField.SelectedItem.Text.ToLower = "category" Then
                    Dim sProductCategory As String = drProduct("ProductCategory").ToString.ToLower
                    If ddlProductFieldValue.SelectedItem.Text.ToLower = sProductCategory Then
                        oFilteredDataTable.ImportRow(dr)
                    End If
                End If
            Next
            oFilteredDataTable.Columns.Remove("LogisticProductKey")
            Return oFilteredDataTable
        Else
            oDataTable.Columns.Remove("LogisticProductKey")
            Return oDataTable
        End If
    End Function

    Private Function sGetSortDirection() As String
        Select Case psItemSortDirection
            Case "ASC"
                psItemSortDirection = "DESC"
            Case "DESC"
                psItemSortDirection = "ASC"
        End Select
        Return psItemSortDirection
    End Function

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

    Protected Sub btnExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExportProductDetails()
    End Sub
    
    Sub ExportProductDetails()
        Dim oDataView As New DataView()
        oDataView = SortItems(dtGetData, isPageIndexChanging:=False)
        If oDataView.Count > 0 Then
    
            Response.Clear()
            'Response.ContentType = "Application/x-msexcel"
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=booked_items.csv")
    
            Dim oDataColumn As DataColumn
            Dim sItem As String
    
            Dim IgnoredItems As New ArrayList
            'IgnoredItems.Add("UserKey")
    
            For Each oDataColumn In oDataView.Table.Columns  ' write column header
                If Not IgnoredItems.Contains(oDataColumn.ColumnName) Then
                    Response.Write(oDataColumn.ColumnName)
                    Response.Write(",")
                End If
            Next
            Response.Write(vbCrLf)
    
            For Each oDataRowView As DataRowView In oDataView
                For Each oDataColumn In oDataView.Table.Columns
                    If Not IgnoredItems.Contains(oDataColumn.ColumnName) Then
                        sItem = (oDataRowView(oDataColumn.ColumnName).ToString)
                        sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                        sItem = ControlChars.Quote & sItem & ControlChars.Quote
                        Response.Write(sItem)
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Else
            lblError.Text = "... no data found"
        End If
    End Sub

    Protected Sub gvItems_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvItems.DataSource = SortItems(gvItems.DataSource, isPageIndexChanging:=True)
        gvItems.PageIndex = e.NewPageIndex
        gvItems.DataBind()
    End Sub
    
    Protected Function SortItems(ByVal dt As DataTable, ByVal isPageIndexChanging As Boolean) As DataView
        If Not dt Is Nothing Then
            Dim dv As New DataView(dt)
            If psItemSortExpression <> String.Empty Then
                If isPageIndexChanging Then
                    dv.Sort = String.Format("{0} {1}", psItemSortExpression, psItemSortDirection)
                Else
                    If gbSortItemChanged Then
                        dv.Sort = String.Format("{0} {1}", psItemSortExpression, "ASC")
                    Else
                        dv.Sort = String.Format("{0} {1}", psItemSortExpression, sGetSortDirection)
                    End If
                End If
            End If
            Return dv
        Else
            Return New DataView
        End If
    End Function
    
    Protected Sub gvItems_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        If psItemSortExpression <> e.SortExpression Then
            gbSortItemChanged = True
            psItemSortDirection = DEFAULT_SORT_DIRECTION
        End If
        psItemSortExpression = e.SortExpression
        Dim nPageIndex As Integer = gvItems.PageIndex
        gvItems.DataSource = SortItems(gvItems.DataSource, isPageIndexChanging:=False)
        gvItems.DataBind()
        gvItems.PageIndex = nPageIndex
    End Sub
    
    Protected Sub gvItems_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.Header Then
            For Each tc As TableCell In gvr.Cells
                Dim lb As LinkButton = tc.Controls(0)
                If lb.Text = psItemSortExpression Then
                    Dim img As New Image
                    If psItemSortDirection = "ASC" Then
                        img.ImageUrl = "~/images/greendown.gif"
                    Else
                        img.ImageUrl = "~/images/orangeup.gif"
                    End If
                    tc.Controls.Add(img)
                End If
            Next
        End If
    End Sub

    Protected Sub btnReselectReportFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectReportFilter()
    End Sub
    
    Protected Sub ReselectReportFilter()
        btnRunReport1.Visible = True
        btnRunReport2.Visible = True
        btnReselectReportFilter1.Visible = False
        btnReselectReportFilter2.Visible = False
        btnExportToExcel1.Visible = False
        btnExportToExcel2.Visible = False
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
        
        psItemSortExpression = DEFAULT_SORT_ITEM
        psItemSortDirection = DEFAULT_SORT_DIRECTION
        gvItems.PageIndex = 0
        
        pnlData.Visible = False
        pbIsDisplayingData = False
        lblError.Text = String.Empty

        ddlProductGroup.Enabled = True

        cbShowConsignmentCost.Enabled = True
        cbFilterByProduct.Enabled = True
        ddlProductField.Enabled = True
        ddlProductFieldValue.Enabled = True

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
            psToDate = tbToDate.Text
            psFromDate = tbFromDate.Text
        Else
            psFromDate = ddlFromDay.SelectedItem.Text & "-" & ddlFromMonth.SelectedItem.Text & "-" & ddlFromYear.SelectedItem.Text
            psToDate = ddlToDay.SelectedItem.Text & "-" & ddlToMonth.SelectedItem.Text & "-" & ddlToYear.SelectedItem.Text
            tbFromDate.Text = psFromDate
            tbToDate.Text = psToDate
        End If
    End Sub
    
    Protected Sub ddlRows_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateItemGridView()
    End Sub

    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
        Call ReselectReportFilter()
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
    
    Private Property psItemSortDirection() As String
        Get
            Return IIf(ViewState("IBH_SortDirection") = Nothing, "ASC", ViewState("IBH_SortDirection"))
        End Get
        Set(ByVal value As String)
            ViewState("IBH_SortDirection") = value
        End Set
    End Property

    Private Property psItemSortExpression() As String
        Get
            Return IIf(ViewState("IBH_SortExpression") = Nothing, DEFAULT_SORT_ITEM, ViewState("IBH_SortExpression"))
        End Get
        Set(ByVal value As String)
            ViewState("IBH_SortExpression") = value
        End Set
    End Property

    Property psToDate() As String
        Get
            Dim o As Object = ViewState("IBH_ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("IBH_ToDate") = Value
        End Set
    End Property
    
    Property psFromDate() As String
        Get
            Dim o As Object = ViewState("IBH_FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("IBH_FromDate") = Value
        End Set
    End Property
    
    Property pbIsDisplayingData() As Boolean
        Get
            Dim o As Object = ViewState("IBH_IsDisplayingData")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IBH_IsDisplayingData") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("IBH_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("IBH_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("IBH_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IBH_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("IBH_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("IBH_IsProductOwner") = Value
        End Set
    End Property
   
    Protected Sub cbShowConsignmentCost_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ReselectReportFilter()
    End Sub
    
    Protected Sub cbFilterByProduct_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            trFilterByProduct.Visible = True
            ddlProductField.SelectedIndex = 0
            If ddlProductFieldValue.SelectedIndex > -1 Then
                ddlProductFieldValue.SelectedIndex = 0
            End If
            ddlProductFieldValue.Visible = False
            lblLegendValue.Visible = False
        Else
            trFilterByProduct.Visible = False
        End If
    End Sub
    
    Protected Sub ddlProductField_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Dim sSQL As String
        If ddl.SelectedIndex > 0 Then
            Dim sField As String = ddl.SelectedItem.Text
            Select Case sField.ToLower
                Case "category"
                    sSQL = "SELECT DISTINCT ProductCategory FROM LogisticProduct WHERE ISNULL(ProductCategory, '') <> '' AND DeletedFlag = 'N' AND CustomerKey = " & Session("CustomerKey")
                    Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
                    ddlProductFieldValue.Items.Clear()
                    For Each dr As DataRow In dt.Rows
                        ddlProductFieldValue.Items.Add(dr(0))
                    Next
                Case Else
            End Select
            ddlProductFieldValue.Visible = True
            lblLegendValue.Visible = True
        Else
            ddlProductFieldValue.Visible = False
            lblLegendValue.Visible = False
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

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Item Booking History Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form id="frmItemBookingHistoryReport" runat="server">
        <table id="tblDateRangeSelector" runat="server" visible="true" width="100%">
            <tr id="Tr1" runat="server">
                <td colspan="4" style="white-space:nowrap">
                  <asp:Label ID="lblPageHeading"
                             runat="server"
                             forecolor="Silver"
                             font-size="Small"
                             font-bold="True"
                             font-names="Arial">Item Booking History Report</asp:Label></td>
            </tr>
            <tr runat="server" visible="true">
                <td style="width: 10%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 25%; white-space: nowrap">
                </td>
                <td style="width: 15%; white-space: nowrap">
                </td>
            </tr>
            <tr runat="server" visible="true" id="trProductGroups">
                <td style="width: 10%; white-space: nowrap">
                    &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                        Visible="False">
                    </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 25%; white-space: nowrap">
                </td>
                <td style="width: 15%; white-space: nowrap">
                    <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                        Text="show product groups" Visible="False" /></td>
            </tr>
            <tr runat="server" visible="true">
                <td style="width: 10%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 25%; white-space: nowrap">
                </td>
                <td style="width: 15%; white-space: nowrap">
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 10%; white-space:nowrap">
                    From:
                        <asp:TextBox ID="tbFromDate"
                                     font-names="Verdana"
                                     font-size="XX-Small"
                                     Width="90"
                                     runat="server">
                          </asp:TextBox>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img alt="" id="Img1"
                                 src="../images/SmallCalendar.gif"
                                 runat="server"
                                 border="0"
                              IE:visible="true"
                                 visible="false"
                               /></a>
                    <span id="spnDateExample1" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2007)</span></td>
                <td style="width: 50%; white-space:nowrap">
                    <span class="informational dark"></span>
                    &nbsp;&nbsp;
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                        <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                            <img alt="" id="Img2" src="../images/SmallCalendar.gif" runat="server" border="0" IE:visible="true" visible="false" />
                         </a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light">(eg&nbsp;12-Jan-2008)</span>
                </td>
                <td style="width: 25%; white-space:nowrap">
                <asp:Button ID="btnRunReport1"
                     runat="server"
                     Text="generate report"
                     Visible="true"
                     OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnExportToExcel1" runat="server" OnClick="btnExportToExcel_Click"
                        Text="export to excel" Visible="False" />
                <asp:Button ID="btnReselectReportFilter1"
                     runat="server"
                     Text="re-select report filter"
                     Visible="false"
                      OnClick="btnReselectReportFilter_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 15%; white-space:nowrap">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between calendar style dates and dropdown style dates"/>
                 </td>
            </tr>
            <tr runat="server" visible="false" id="DropdownInterface">
                <td style="width: 10%; white-space:nowrap">
                    From: &nbsp;<asp:DropDownList ID="ddlFromDay" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                    </asp:DropDownList></td>
                <td style="width: 50%; white-space:nowrap">
                    <span class="informational dark"></span>
                    &nbsp;
                    &nbsp;&nbsp;
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
                    </asp:DropDownList>&nbsp;
                </td>
                <td style="width: 25%; white-space:nowrap">
                <asp:Button ID="btnRunReport2"
                     runat="server"
                     Text="generate report"
                      OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnExportToExcel2" runat="server" OnClick="btnExportToExcel_Click"
                        Text="export to excel" Visible="False" />
                    <asp:Button ID="btnReselectReportFilter2"
                     runat="server"
                     Text="re-select report filter"
                     Visible="false"
                      OnClick="btnReselectReportFilter_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 15%; white-space:nowrap">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface"/>
                 </td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td>
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate" ErrorMessage=" - invalid format for date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage=" - year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td>
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate" ErrorMessage=" - invalid format for date - use dd-mmm-yy"
                        Font-Names="Verdana" Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True" ErrorMessage=" - year before 2000, after 2020, or not a valid date!"
                            Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                        Font-Size="XX-Small" ForeColor="Red"></asp:Label></td>
                <td>
                    <asp:CheckBox ID="cbShowConsignmentCost" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        Text="show consignment cost" AutoPostBack="True" OnCheckedChanged="cbShowConsignmentCost_CheckedChanged" />
                    &nbsp;<asp:CheckBox ID="cbFilterByProduct" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" oncheckedchanged="cbFilterByProduct_CheckedChanged" Text="filter by product" />
                </td>
                <td>
                </td>
            </tr>
            <tr runat="server">
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr runat="server" visible="false" id="trFilterByProduct">
                <td align="right">
                    Filter by:</td>
                <td>
                    <asp:DropDownList ID="ddlProductField" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onselectedindexchanged="ddlProductField_SelectedIndexChanged" AutoPostBack="True">
                        <asp:ListItem Value="0">- please select -</asp:ListItem>
                        <asp:ListItem>Category</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td align="right">
                    <asp:Label ID="lblLegendValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Value:"></asp:Label>
&nbsp;<asp:DropDownList ID="ddlProductFieldValue" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                    </asp:DropDownList>
                &nbsp;
                </td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
        <asp:Panel id="pnlData" runat="server" Visible="false" Width="100%">
            <asp:GridView ID="gvItems" runat="server" Width="100%" OnPageIndexChanging="gvItems_PageIndexChanging" OnSorting="gvItems_Sorting" AllowPaging="True" AllowSorting="True" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" OnRowDataBound="gvItems_RowDataBound">
                <PagerStyle Font-Bold="False" Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
                <AlternatingRowStyle BackColor="WhiteSmoke" />
                <PagerSettings Position="TopAndBottom" />
            </asp:GridView>
            <br />
            <table>
                <tr>
                    <td style="width: 300px">
                        &nbsp;<asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green" Visible="false"></asp:Label></td>
                    <td style="width: 500px">
                        Items per page:
                        <asp:DropDownList ID="ddlRows" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlRows_SelectedIndexChanged">
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                            <asp:ListItem>100</asp:ListItem>
                        </asp:DropDownList></td>
                </tr>
            </table>
        </asp:Panel>
        <br />
        <asp:Label id="lblError" runat="server" Font-Names="Arial" Font-Size="XX-Small" ForeColor="red"></asp:Label>
    </form>
</body>
</html>