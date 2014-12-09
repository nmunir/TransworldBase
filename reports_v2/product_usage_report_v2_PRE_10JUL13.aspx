<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<%@ Register TagPrefix="telerik" Namespace="Telerik.Charting" Assembly="Telerik.Web.UI" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>

<script runat="server">

    ' Product Usage Report - HEADERLESS VERSION FOR COMMON REPORTING FACILITY - CN
    ' Shows total Goods Out for all products over selected period
    ' TO DO
    ' Write lean sproc for retrieving export data
    ' Write sproc to get product name / value date for export data
    ' handle initialisation of From and To dat properties in either date format in Page_Load (not sure if this is required)
    ' Check dates are in correct format going to database (UK/US problem)

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Dim iDefaultHistoryPeriod As Integer = -12      'last 12 months

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
            
            ' ALSO NEED TO COPE WITH EITHER DATE FORMAT  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            
            psFromDate = tbFromDate.Text
            psToDate = tbToDate.Text
            'WebChart1.Visible = False
            RadChart1.Visible = False
            Call SelectDateRange()
        End If
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
    
    Protected Sub btnReselectDateRange_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SelectDateRange()
    End Sub
    
    Protected Sub SelectDateRange()
        pnlData.Visible = False
        btnRunReport1.Visible = True
        btnRunReport2.Visible = True
        btnReselectDateRange1.Visible = False
        btnReselectDateRange2.Visible = False
        tbFromDate.Enabled = True
        tbToDate.Enabled = True
        ddlFromDay.Enabled = True
        ddlFromMonth.Enabled = True
        ddlFromYear.Enabled = True
        ddlToDay.Enabled = True
        ddlToMonth.Enabled = True
        ddlToYear.Enabled = True
        lblDateExample1.Visible = True
        lblDateExample2.Visible = True
        imgCalendarButton1.Visible = True
        imgCalendarButton2.Visible = True
        cbShowProductsWithMovements.Enabled = True
        cbShowProductsWithNoMovements.Enabled = True
        lblLegendChartBy.Visible = False
        rbChartPeriodWeek.Visible = False
        rbChartPeriodMonth.Visible = False
        btnExportToExcel.Visible = False

        cbShowProductsWithNoMovements.Visible = True
        cbShowProductsWithMovements.Visible = True
        cbIncludeArchivedProducts.Visible = True
       
        lblReportInclusion.Text = String.Empty	 
        ddlProductGroup.Enabled = True
    End Sub

    Protected Sub ShowDataPanel()
        pnlData.Visible = True
    End Sub
    
    Protected Sub GetDateRange()
        btnRunReport1.Visible = False
        btnRunReport2.Visible = False
        btnReselectDateRange1.Visible = True
        btnReselectDateRange2.Visible = True
        tbFromDate.Enabled = False
        tbToDate.Enabled = False
        ddlFromDay.Enabled = False
        ddlFromMonth.Enabled = False
        ddlFromYear.Enabled = False
        ddlToDay.Enabled = False
        ddlToMonth.Enabled = False
        ddlToYear.Enabled = False
        cbShowProductsWithMovements.Enabled = False
        cbShowProductsWithNoMovements.Enabled = False
        
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
            lblDateExample1.Visible = False
            lblDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            
            lblReportGeneratedDateTime.Visible = True
            Call BindProductGrid("ProductCode")
            Call ShowDataPanel()
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
        lblLegendChartBy.Visible = True
        rbChartPeriodWeek.Visible = True
        rbChartPeriodMonth.Visible = True

        btnExportToExcel.Visible = True
       
        If cbShowProductsWithNoMovements.Checked AndAlso cbShowProductsWithMovements.Checked Then
            If cbIncludeArchivedProducts.Checked Then
                lblReportInclusion.Text = " - all products, including archived products"
            Else
                lblReportInclusion.Text = " - all unarchived products"
            End If
        ElseIf cbShowProductsWithNoMovements.Checked Then
            If cbIncludeArchivedProducts.Checked Then
                lblReportInclusion.Text = " - products, including archived products, with no movements"
            Else
                lblReportInclusion.Text = " - unarchived products with no movements"
            End If
        Else
            If cbIncludeArchivedProducts.Checked Then
                lblReportInclusion.Text = " - products, including archived products, with movements"
            Else
                lblReportInclusion.Text = " - unarchived products with movements"
            End If
        End If

        cbShowProductsWithNoMovements.Visible = False
        cbShowProductsWithMovements.Visible = False
        cbIncludeArchivedProducts.Visible = False
    End Sub
    
    Protected Sub BindProductGrid(ByVal sSortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spReport_Product_Usage6", oConn)
        Dim nFlag As Integer = 0
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = CDate(psFromDate)
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = DateAdd("D", 1, CDate(psToDate))

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        If cbShowProductsWithNoMovements.Checked Then
            nFlag += 1
        End If
        If cbShowProductsWithMovements.Checked Then
            nFlag += 2
        End If
        
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Flag", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@Flag").Value = nFlag - 1

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SortKey", SqlDbType.VarChar, 50))
        oAdapter.SelectCommand.Parameters("@SortKey").Value = sSortField

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ArchiveFlag", SqlDbType.Int))
        If cbIncludeArchivedProducts.Checked Then
            oAdapter.SelectCommand.Parameters("@ArchiveFlag").Value = 1
        Else
            oAdapter.SelectCommand.Parameters("@ArchiveFlag").Value = 0
        End If

        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                dgrdProducts.DataSource = oDataTable
                dgrdProducts.DataBind()
                lblError.Text = ""
                lblReportGeneratedDateTime.Visible = True
            Else
                lblError.Text = "no data found for this date range"
                lblReportGeneratedDateTime.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Sub SortProductColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        BindProductGrid(E.SortExpression)
    End Sub

    Protected Sub lnkbtnToggleSelectionStyle_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CalendarInterface.Visible = True Then
            CalendarInterface.Visible = False
            DropdownInterface.Visible = True
        Else
            CalendarInterface.Visible = True
            DropdownInterface.Visible = False
        End If
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
    End Sub

    Protected Sub cbShowProductsWithNoMovements_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cbShowProductsWithNoMovements As CheckBox = sender
        If Not cbShowProductsWithNoMovements.Checked Then
            If Not cbShowProductsWithMovements.Checked Then
                cbShowProductsWithMovements.Checked = True
            End If
        End If
        Call SelectDateRange()
    End Sub

    Protected Sub cbShowProductsWithMovements_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cbShowProductsWithMovements As CheckBox = sender
        If Not cbShowProductsWithMovements.Checked Then
            If Not cbShowProductsWithNoMovements.Checked Then
                cbShowProductsWithNoMovements.Checked = True
            End If
        End If
        Call SelectDateRange()
    End Sub
    
    Protected Sub btnExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExportToExcel()
    End Sub
    
    Protected Sub ExportToExcel()
        Dim oDataTable As DataTable
        Dim sCSVString As String
        oDataTable = GenerateExportData()
        sCSVString = ConvertDataTableToCSVString(oDataTable)
        Call ExportCSVData(sCSVString)
    End Sub
    
    Protected Function GenerateExportData() As DataTable
        Dim oProductUsageByPeriodDataRow As DataRow
        Dim bTableConstructed As Boolean = False
        Dim dateStartSegment As Date
        Dim dateFinishSegment As Date
        Dim dtProductUsageByPeriod As New DataTable
        Dim dtSegment As DataTable
        dateStartSegment = CDate(psFromDate)
        dateFinishSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
        
        dtSegment = GetSegment(dateStartSegment, dateFinishSegment)
        If Not bTableConstructed Then
            dtProductUsageByPeriod.Columns.Add(New DataColumn("Period").ToString, GetType(String))
            For Each drSegment As DataRow In dtSegment.Rows
                dtProductUsageByPeriod.Columns.Add(New DataColumn(drSegment("LogisticProductKey").ToString, GetType(Integer)))
            Next
            bTableConstructed = True
        End If
        
        While dateFinishSegment < CDate(psToDate)
            oProductUsageByPeriodDataRow = dtProductUsageByPeriod.NewRow
            oProductUsageByPeriodDataRow("Period") = dateStartSegment.ToShortDateString
            For Each drSegment As DataRow In dtSegment.Rows
                oProductUsageByPeriodDataRow(drSegment("LogisticProductKey").ToString) = drSegment("Quantity")   ' one of the few cases where ToString is actually necessary since the compiler does not know to do it implicitly
            Next
            dtProductUsageByPeriod.Rows.Add(oProductUsageByPeriodDataRow)
            dateStartSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
            dateFinishSegment = DateAdd(DateInterval.Day, 7, dateFinishSegment)
            dtSegment = GetSegment(dateStartSegment, dateFinishSegment)
        End While
        GenerateExportData = dtProductUsageByPeriod
    End Function
    
    Protected Function GetSegment(ByVal dateStartSegment As Date, ByVal dateFinishSegment As Date) As DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spReport_Product_UsageExport2", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = dateStartSegment.ToShortDateString
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = dateFinishSegment.ToShortDateString

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        GetSegment = oDataTable
    End Function
    
    Protected Function GenerateChartData(ByVal nLogisticProductKey As Integer) As Dictionary(Of String, Integer)
        Dim dicChartData As New Dictionary(Of String, Integer)
        Dim dateStartSegment As Date
        Dim dateFinishSegment As Date
        Dim alChartData As New ArrayList
        Dim alChartDates As New ArrayList
        dateStartSegment = CDate(psFromDate)
        dateFinishSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
        alChartData.Add(GetSegmentForSingleProduct(nLogisticProductKey, dateStartSegment, dateFinishSegment))
        While dateFinishSegment < CDate(psToDate)
            dateStartSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
            dateFinishSegment = DateAdd(DateInterval.Day, 7, dateFinishSegment)
            dicChartData.Add(dateStartSegment.ToString("dd-MMM-yyyy"), GetSegmentForSingleProduct(nLogisticProductKey, dateStartSegment, dateFinishSegment))
        End While
        GenerateChartData = dicChartData
    End Function
    
    Protected Function GenerateChartData2(ByVal nLogisticProductKey As Integer) As List(Of Product)
        'Dim dicChartData As New Dictionary(Of String, Integer)
        Dim sDateFormat As String
        Dim lstProducts As New List(Of Product)
        Dim dateStartSegment As Date
        Dim dateFinishSegment As Date
        Dim alChartData As New ArrayList
        Dim alChartDates As New ArrayList
        dateStartSegment = CDate(psFromDate)
        If rbChartPeriodWeek.Checked Then
            dateFinishSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
        Else
            dateFinishSegment = DateAdd(DateInterval.Month, 1, dateStartSegment)
        End If
        alChartData.Add(GetSegmentForSingleProduct(nLogisticProductKey, dateStartSegment, dateFinishSegment))
        While dateFinishSegment < CDate(psToDate)
            If rbChartPeriodWeek.Checked Then
                dateStartSegment = DateAdd(DateInterval.Day, 7, dateStartSegment)
                dateFinishSegment = DateAdd(DateInterval.Day, 7, dateFinishSegment)
                sDateFormat = "dd-MMM-yy"
            Else
                dateStartSegment = DateAdd(DateInterval.Month, 1, dateStartSegment)
                dateFinishSegment = DateAdd(DateInterval.Month, 1, dateFinishSegment)
                sDateFormat = "MMM-yy"
            End If
            lstProducts.Add(New Product(GetSegmentForSingleProduct(nLogisticProductKey, dateStartSegment, dateFinishSegment), dateStartSegment.ToString(sDateFormat)))
        End While
        GenerateChartData2 = lstProducts
    End Function

    Public Class Product
        Public Sub New(ByVal usage As Integer, ByVal period As String)
            _usage = usage
            _period = period
        End Sub
        Private _usage As Integer
        Public Property usage() As Integer
            Get
                Return _usage
            End Get
            Set(ByVal value As Integer)
                _usage = value
            End Set
        End Property
        Private _period As String
        Public Property period() As String
            Get
                Return _period
            End Get
            Set(ByVal value As String)
                _period = value
            End Set
        End Property
    End Class
   
    Protected Function GetSegmentForSingleProduct(ByVal LogisticProductKey As Integer, ByVal dateStartSegment As Date, ByVal dateFinishSegment As Date) As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spReport_Product_UsageChartSingleItem", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = LogisticProductKey
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = dateStartSegment.ToString("dd-MMM-yyyy")
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = dateFinishSegment.ToString("dd-MMM-yyyy")

        Try
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        If oDataTable.Rows.Count = 0 Then
            GetSegmentForSingleProduct = 0
        Else
            GetSegmentForSingleProduct = oDataTable.Rows(0).Item("Quantity")
        End If
    End Function
    
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "ProductUsageData.csv")
        'Response.ContentType = "application/vnd.ms-excel"
        Response.ContentType = "text/csv"
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()

        ' Stop execution of the current page
        Response.End()
    End Sub
    
    Public Function GetProductIdFromKey(ByVal nLogisticProductKey As Integer) As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spReport_Product_UsageGetProductIdFromKey", oConn)

        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LogisticProductKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@LogisticProductKey").Value = nLogisticProductKey
    
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
        GetProductIdFromKey = oDataTable.Rows(0).Item(0)
    End Function
    
    Public Function ConvertDataTableToCSVString(ByVal oDataTable As DataTable) As String
        Dim sbResult As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRow

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            If IsNumeric(oDataColumn.ColumnName) Then
                sbResult.Append(GetProductIdFromKey(oDataColumn.ColumnName))
            Else
                sbResult.Append(oDataColumn.ColumnName)
            End If
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            sbResult.Length = sbResult.Length - 1
        End If
        sbResult.Append(Environment.NewLine)
    
        For Each oDataRow In oDataTable.Rows
            For Each oDataColumn In oDataTable.Columns
                sbResult.Append(oDataRow(Replace(oDataColumn.ColumnName, ",", " ")))  ' replace any commas with spaces
                sbResult.Append(",")
            Next oDataColumn
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
    
    Protected Sub lnkbtnChart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton
        lb = sender
        lblLegendChartBy.Visible = False
        rbChartPeriodWeek.Visible = False
        rbChartPeriodMonth.Visible = False
        btnExportToExcel.Visible = False
        Call GenerateChart(CInt(lb.CommandArgument))
    End Sub

    Protected Sub GenerateChart(ByVal nLogisticProductKey As Integer)
        Dim lstChartData As List(Of Product) = GenerateChartData2(nLogisticProductKey)
        RadChart1.DataSource = lstChartData
        RadChart1.Series(0).DataYColumn = "usage"
        RadChart1.PlotArea.XAxis.DataLabelsColumn = "period"
        If lstChartData.Count <= 12 AndAlso rbChartPeriodMonth.Checked Then
            RadChart1.SeriesOrientation = ChartSeriesOrientation.Vertical
        Else
            RadChart1.SeriesOrientation = ChartSeriesOrientation.Horizontal
            RadChart1.Height = lstChartData.Count * 20
        End If
        RadChart1.DataBind()
        RadChart1.Visible = True
        RadChart1.AutoLayout = True
        pnlChart.Visible = True
        pnlData.Visible = False
        Dim sProductId As String = GetProductIdFromKey(nLogisticProductKey)
        RadChart1.ChartTitle.TextBlock.Text = sProductId & " - " & psFromDate & " to " & psToDate
        lblChart.Text = "Product usage for " & sProductId & " from " & psFromDate & " to " & psToDate
    End Sub

    Protected Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlChart.Visible = False
        pnlData.Visible = True
        RadChart1.Visible = False
        lblLegendChartBy.Visible = True
        rbChartPeriodWeek.Visible = True
        rbChartPeriodMonth.Visible = True
        btnExportToExcel.Visible = True
    End Sub
    
    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
    End Sub

    Protected Sub ShowProductGroups()
        ddlProductGroup.Visible = True
        Call PopulateProductGroups(0)
        btnShowProductGroups.Visible = False
    End Sub
    
    Protected Sub ddlProductGroup_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.Items(0).Value = -1 Then
            ddlProductGroup.Items.RemoveAt(0)
        End If
        pnSelectedProductGroup = ddl.SelectedValue
        btnRunReport1.Enabled = True
        btnRunReport2.Enabled = True
    End Sub
    
    Property psToDate() As String
        Get
            Dim o As Object = ViewState("PUR_ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PUR_ToDate") = Value
        End Set
    End Property
    
    Property psFromDate() As String
        Get
            Dim o As Object = ViewState("PUR_FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("PUR_FromDate") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("PUR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("PUR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("PUR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PUR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("PUR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PUR_IsProductOwner") = Value
        End Set
    End Property
   
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Product Usage Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" /Product Usage Report</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form id="frmReport" runat="server">
        <table>
            <tr id="Tr1" runat="server" visible="true">
                <td colspan="4" style="white-space: nowrap">
                    <asp:Label ID="lblPageHeading" runat="server" ForeColor="Silver" Font-Size="Small" Font-Bold="True" Font-Names="Verdana" Text="Product Usage Report"/><asp:Label ID="lblReportInclusion" runat="server" Font-Names="Verdana" Font-Size="Small" ForeColor="Silver"/>

                    <br />
                    <br />
                </td>
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
            <tr runat="server" visible="true" id="trProductGroups">
                <td colspan="2" style="white-space: nowrap; height: 14px">
                    &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                        Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                        Visible="False">
                    </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Bold="True"
                        Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                <td style="width: 253px; height: 14px">
                </td>
                <td style="width: 169px; height: 14px">
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
                <td style="width: 265px; white-space: nowrap">
                    <span class="informational dark"><asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="From:"></asp:Label></span>
                    <asp:TextBox ID="tbFromDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90" runat="server">
                    </asp:TextBox>
                    <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img1" alt="" src="../images/SmallCalendar.gif" runat="server" border="0"
                            ie:visible="true" visible="false" /></a>
                            <asp:Label ID="lblDateExample1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="(eg 12-Jan-09)"/>
                </td>
                <td style="width: 265px; white-space: nowrap">
                    <span class="informational dark"><asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="To:"/></span>
                    <asp:TextBox ID="tbToDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90" runat="server">
                    </asp:TextBox>
                    <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                        <img id="Img2" alt="" src="../images/SmallCalendar.gif" runat="server" border="0"
                            ie:visible="true" visible="false" /></a>
                         <asp:Label ID="lblDateExample2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="(eg 12-Jan-10)"/>
                </td>
                <td style="width: 253px">
                    <asp:Button ID="btnRunReport1" runat="server" Text="generate report" Visible="true"
                        OnClick="btnRunReport_Click" Font-Names="Verdana" Font-Size="XX-Small" Width="170px" />
                    <asp:Button ID="btnReselectDateRange1" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectDateRange_Click" Font-Names="Verdana" Font-Size="XX-Small" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between calendar interface and dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton></td>
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
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromMonth" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
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
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlFromYear" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
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
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
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
                    </asp:DropDownList>&nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
                    </asp:DropDownList>&nbsp;</td>
                <td style="width: 253px; height: 26px;">
                    <asp:Button ID="btnRunReport2" runat="server" Text="generate report" OnClick="btnRunReport_Click"
                        Font-Names="Verdana" Font-Size="XX-Small" Width="170px" />
                    <asp:Button ID="btnReselectDateRange2" runat="server" Text="re-select report filter"
                        Visible="false" OnClick="btnReselectDateRange_Click" Font-Names="Verdana" Font-Size="XX-Small" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px; height: 26px;">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                        ToolTip="toggles between calendar interface and dropdown interface">change&nbsp;selection&nbsp;style</asp:LinkButton></td>
            </tr>
            <tr id="Tr2" runat="server" visible="true">
                <td colspan="4" style="white-space: nowrap">
                    <br />
                    <asp:CheckBox ID="cbShowProductsWithNoMovements" runat="server" Checked="True" OnCheckedChanged="cbShowProductsWithNoMovements_CheckedChanged"
                        Text="include products with no movements" AutoPostBack="True" />
                    &nbsp;
                    <asp:CheckBox ID="cbShowProductsWithMovements" runat="server" Checked="True" Text="include products with movements"
                        OnCheckedChanged="cbShowProductsWithMovements_CheckedChanged" AutoPostBack="True" />
                    <asp:Label ID="lblLegendChartBy" runat="server" Text="Chart by:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
                    <asp:RadioButton ID="rbChartPeriodWeek" runat="server" GroupName="ChartPeriod" Text="week" Font-Names="Verdana" Font-Size="XX-Small" />
                    <asp:RadioButton ID="rbChartPeriodMonth" runat="server" Checked="True" GroupName="ChartPeriod" Text="month" Font-Names="Verdana" Font-Size="XX-Small" />

                    &nbsp; &nbsp;
                    <asp:Button ID="btnExportToExcel" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        OnClick="btnExportToExcel_Click" Text="export time series to excel" /><br />
                </td>
            </tr>
            <tr id="Tr3" runat="server" visible="true">
                <td colspan="4" style="white-space: nowrap">
                    <asp:CheckBox ID="cbIncludeArchivedProducts" runat="server" 
                        Text="include archived products" />
                </td>
            </tr>
            <tr runat="server" visible="true" id="DateValidationMessages">
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate"
                        ErrorMessage="invalid format for date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                    <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                        CultureInvariantValues="True" ErrorMessage="year before 2000, after 2020, or not a valid date!"
                        Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1"
                        ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red"></asp:Label></td>
                <td style="width: 265px">
                    <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate"
                        ErrorMessage="invalid format for date - use dd-mmm-yy" Font-Names="Verdana"
                        Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                        ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                            ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True"
                            ErrorMessage="year before 2000, after 2020, or not a valid date!" Font-Names="Verdana"
                            Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface"
                            EnableClientScript="False" Type="Date"></asp:RangeValidator>
                    <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        ForeColor="Red"></asp:Label></td>
                <td style="width: 253px">
                </td>
                <td style="width: 169px">
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlData" runat="server" Width="100%">
            <asp:DataGrid ID="dgrdProducts" runat="server" Width="95%" Font-Names="Arial" Font-Size="XX-Small" CellSpacing="7" AutoGenerateColumns="False" GridLines="None" AllowSorting="True"
                OnSortCommand="SortProductColumns" AlternatingItemStyle-BackColor="Gray">
                <HeaderStyle Font-Size="XX-Small" Font-Names="Verdana" ForeColor="Blue"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticProductKey" SortExpression="LogisticProductKey"
                        HeaderText="No." DataFormatString="{0:000000}">
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Product Code">
                        <ItemStyle Wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDate" SortExpression="ProductDate" HeaderText="Product Date">
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription"
                        HeaderText="Description">
                        <ItemStyle Wrap="True"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Quantity" SortExpression="Quantity" HeaderText="Quantity"
                        DataFormatString="{0:#,##0}">
                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                        <ItemStyle Wrap="False" HorizontalAlign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemTemplate>
                            <asp:LinkButton ID="lnkbtnChart" OnClick="lnkbtnChart_Click" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"LogisticProductKey") %>'
                                runat="server">chart</asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <AlternatingItemStyle BackColor="#DDDDDD" />
            </asp:DataGrid>
            <br />
            <asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" Font-Size="XX-Small"
                Font-Names="Verdana, Sans-Serif" ForeColor="Green" Visible="false"></asp:Label>
        </asp:Panel>
        <asp:Panel ID="pnlChart" runat="server" Visible="false" Width="100%">
            <table style="width: 100%">
                <tr>
                    <td style="width: 50%">
                        <asp:Label ID="lblChart" runat="server"></asp:Label></td>
                    <td style="width: 50%" align="right">
                        <asp:Button ID="btnBack" runat="server" Text="back" OnClick="btnBack_Click" /></td>
                </tr>
                <tr>
                    <td align="center" colspan="2">
                        &nbsp;<telerik:RadChart ID="RadChart1" runat="server" Width="700px">
                            <Series>
                                <telerik:ChartSeries Name="Product Usage">
                                </telerik:ChartSeries>
                            </Series>
                        </telerik:RadChart>
                    </td>
                </tr>
            </table>
            <br />
            &nbsp;</asp:Panel>
        <br />
        <asp:Label ID="lblError" runat="server" Font-Names="Arial" Font-Size="XX-Small" ForeColor="red"></asp:Label>
    </form>
</body>
</html>
