<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private Shared _exportData As DataView
    
    Private Const C_HTTP_HEADER_CONTENT As String = "Content-Disposition"
    Private Const C_HTTP_ATTACHMENT As String = "attachment;filename="
    Private Const C_HTTP_INLINE As String = "inline;filename="
    Private Const C_HTTP_CONTENT_TYPE_OCTET As String = "application/octet-stream"
    Private Const C_HTTP_CONTENT_TYPE_EXCEL As String = "application/ms-excel"
    Private Const C_HTTP_CONTENT_LENGTH As String = "Content-Length"
    Private Const C_QUERY_PARAM_CRITERIA As String = "Criteria"
    Private Const C_ERROR_NO_RESULT As String = "Data not found"
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
        If Not IsPostBack Then
            Dim iThisDay As Integer = Day(Now)
            Dim iThisMonth As Integer = DatePart(DateInterval.Month, Now)
            Dim iThisYear As Integer = Year(Now)
            Response.Cache.SetCacheability(System.Web.HttpCacheability.NoCache)

            ddlFromYear.Items.Add(iThisYear - 3)
            ddlFromYear.Items.Add(iThisYear - 2)
            ddlFromYear.Items.Add(iThisYear - 1)
            ddlFromYear.Items.Add(iThisYear)
            ddlFromYear.Items.Add(iThisYear + 1)
    
            ddlToYear.Items.Add(iThisYear - 3)
            ddlToYear.Items.Add(iThisYear - 2)
            ddlToYear.Items.Add(iThisYear - 1)
            ddlToYear.Items.Add(iThisYear)
            ddlToYear.Items.Add(iThisYear + 1)
    
            ddlFromDay.SelectedIndex = iThisDay
            ddlFromMonth.SelectedIndex = iThisMonth - 1
            ddlFromYear.SelectedIndex = 3
    
            ddlToDay.SelectedIndex = iThisDay
            ddlToMonth.SelectedIndex = iThisMonth
            ddlToYear.SelectedIndex = 3
    
            Call LoadCostCentres()
            Call LoadProductCodes()
            Call LoadCategories()
            Call LoadSubCategories()
            Call ShowReportCriteria()
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
            btnRunCostCentreReport.Enabled = False
            btnRunProductCodeReport.Enabled = False
            btnRunCategoryReport.Enabled = False
            btnRunSubCategoryReport.Enabled = False
            pnSelectedProductGroup = -1
        End If
    End Sub
    
    Protected Sub btn_Run_CostCentre_Report_Click(ByVal s As Object, ByVal e As System.EventArgs)
        If ValidDate() Then
            lblReportTitle.Text = "All consignments where Cost Centre = '" & drop_CostCentre.SelectedItem.Text & "' between " & psFromDate & " and " & psToDate
            psReportName = "CostCentreReport"
            Call RunReport("ProductCode")
            Call ShowReportData()
        End If
    End Sub
    
    Protected Sub btn_Run_ProductCode_Report_Click(ByVal s As Object, ByVal e As System.EventArgs)
        If ValidDate() Then
            lblReportTitle.Text = "All consignments where Product Code = '" & ddlProductCode.SelectedItem.Text & "' between " & psFromDate & " and " & psToDate
            psReportName = "ProductCodeReport"
            Call RunReport("ProductCode")
            Call ShowReportData()
        End If
    End Sub
    
    Protected Sub btn_Run_Category_Report_Click(ByVal s As Object, ByVal e As System.EventArgs)
        If ValidDate() Then
            lblReportTitle.Text = "All consignments where Category = '" & ddlCategory.SelectedItem.Text & "' between " & psFromDate & " and " & psToDate
            psReportName = "CategoryReport"
            Call RunReport("ProductCode")
            Call ShowReportData()
        End If
    End Sub
    
    Protected Sub btn_Run_SubCategory_Report_Click(ByVal s As Object, ByVal e As System.EventArgs)
        If ValidDate() Then
            lblReportTitle.Text = "All consignments where Sub Category = '" & ddlSubCategory.SelectedItem.Text & "' between " & psFromDate & " and " & psToDate
            psReportName = "SubCategoryReport"
            Call RunReport("ProductCode")
            Call ShowReportData()
        End If
    End Sub
    
    Protected Sub btn_DownloadCSVFile_Click(ByVal s As Object, ByVal e As EventArgs)
        Call ExportCSVData()
    End Sub
    
    Protected Sub SortReportColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        Call RunReport(E.SortExpression)
    End Sub
    
    Protected Sub ShowReportCriteria()
        pnlReportCriteria.Visible = True
        pnlReportData.Visible = False
    End Sub
    
    Protected Sub ShowReportData()
        pnlReportCriteria.Visible = True
        pnlReportData.Visible = True
    End Sub
    
    Protected Sub LoadCostCentres()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetCostCentres2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int, 4)
        oParam.Value = CLng(Session("CustomerKey"))

        Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@ProductGroup", SqlDbType.Int)
        oParam2.Value = pnSelectedProductGroup
        
        Try
            oConn.Open()
            Dim oReader As SqlDataReader = oCmd.ExecuteReader()
            drop_CostCentre.DataSource = oReader
            drop_CostCentre.DataValueField = "ProductDepartmentId"
            drop_CostCentre.DataBind()
            oReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub LoadProductCodes()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetCodes2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int)
        oParam.Value = CLng(Session("CustomerKey"))

        Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@ProductGroup", SqlDbType.Int)
        oParam2.Value = pnSelectedProductGroup
        
        Try
            oConn.Open()
            Dim oReader As SqlDataReader = oCmd.ExecuteReader()
            ddlProductCode.DataSource = oReader
            ddlProductCode.DataValueField = "ProductCode"
            ddlProductCode.DataBind()
            oReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub LoadCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetCategories2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int)
        oParam.Value = CLng(Session("CustomerKey"))

        Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@ProductGroup", SqlDbType.Int)
        oParam2.Value = pnSelectedProductGroup
        
        Try
            oConn.Open()
            Dim oReader As SqlDataReader = oCmd.ExecuteReader()
            ddlCategory.DataSource = oReader
            ddlCategory.DataValueField = "Category"
            ddlCategory.DataBind()
            oReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub LoadSubCategories()
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As New SqlCommand("spASPNET_Product_GetDistinctSubCategories2", oConn)
        oCmd.CommandType = CommandType.StoredProcedure

        Dim oParam As SqlParameter = oCmd.Parameters.Add("@CustomerKey", SqlDbType.Int)
        oParam.Value = CLng(Session("CustomerKey"))

        Dim oParam2 As SqlParameter = oCmd.Parameters.Add("@ProductGroup", SqlDbType.Int)
        oParam2.Value = pnSelectedProductGroup
        
        Try
            oConn.Open()
            Dim oReader As SqlDataReader = oCmd.ExecuteReader()
            ddlSubCategory.DataSource = oReader
            ddlSubCategory.DataValueField = "SubCategory"
            ddlSubCategory.DataBind()
            oReader.Close()
        Catch ex As SqlException
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Sub RunReport(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter
        Select Case psReportName
            Case "CostCentreReport"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Report_GetConsignments1", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = psFromDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = psToDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CostCentre", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@CostCentre").Value = drop_CostCentre.SelectedItem.Text
            Case "ProductCodeReport"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Report_GetConsignments2", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = psFromDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = psToDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductCode", SqlDbType.NVarChar, 25))
                oAdapter.SelectCommand.Parameters("@ProductCode").Value = ddlProductCode.SelectedItem.Text
            Case "CategoryReport"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Report_GetConsignments3", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = psFromDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = psToDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Category", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@Category").Value = ddlCategory.SelectedItem.Text
            Case "SubCategoryReport"
                oAdapter.SelectCommand = New SqlCommand("spASPNET_Report_GetConsignments4", oConn)
                oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int, 4))
                oAdapter.SelectCommand.Parameters("@CustomerKey").Value = CLng(Session("CustomerKey"))
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@FromDate").Value = psFromDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
                oAdapter.SelectCommand.Parameters("@ToDate").Value = psToDate
                oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SubCategory", SqlDbType.NVarChar, 50))
                oAdapter.SelectCommand.Parameters("@SubCategory").Value = ddlSubCategory.SelectedItem.Text
        End Select
        Try
            oAdapter.Fill(oDataSet, "Consignments")
            pdvReportDataDataView = oDataSet.Tables("Consignments").DefaultView
            pdvReportDataDataView.Sort = SortField
            If pdvReportDataDataView.Count > 0 Then
                lblError.Text = ""
                grid_ReportData.DataSource = pdvReportDataDataView
                grid_ReportData.DataBind()
                grid_ReportData.Visible = True
            Else
                lblError.Text = "No records found"
                grid_ReportData.Visible = False
            End If
    
        Catch ex As SqlException
            grid_ReportData.Visible = False
            lblError.Text = ex.ToString
        End Try
        oConn.Close()
    End Sub
    
    Protected Function ValidDate() As Boolean
        Dim sFromDayPart As String = String.Empty
        Dim sFromMonthPart As String = String.Empty
        Dim sFromYearPart As String = String.Empty
        Dim sToDayPart As String = String.Empty
        Dim sToMonthPart As String = String.Empty
        Dim sToYearPart As String = String.Empty
        Dim sMessage As String = String.Empty
        If ddlFromDay.SelectedItem.Text = "DAY" Then
            ValidDate = False
            sMessage = "[FROM DAY]"
        End If
        If ddlFromMonth.SelectedItem.Text = "MONTH" Then
            ValidDate = False
            sMessage &= "[FROM MONTH]"
        End If
        If ddlFromYear.SelectedItem.Text = "YEAR" Then
            ValidDate = False
            sMessage &= "[FROM YEAR]"
        End If
        If ddlToDay.SelectedItem.Text = "DAY" Then
            ValidDate = False
            sMessage &= "[TO DAY]"
        End If
        If ddlToMonth.SelectedItem.Text = "MONTH" Then
            ValidDate = False
            sMessage &= "[TO MONTH]"
        End If
        If ddlToYear.SelectedItem.Text = "YEAR" Then
            ValidDate = False
            sMessage &= "[TO YEAR]"
        End If
    
        If sMessage <> "" Then
            lblDateError.Text = "Invalid date: " & sMessage
        Else
            ValidDate = True
            lblDateError.Text = ""
            sFromDayPart = ddlFromDay.SelectedItem.Text
            sFromMonthPart = ddlFromMonth.SelectedItem.Text
            sFromYearPart = ddlFromYear.SelectedItem.Text
            sToDayPart = ddlToDay.SelectedItem.Text
            sToMonthPart = ddlToMonth.SelectedItem.Text
            sToYearPart = ddlToYear.SelectedItem.Text
            Try
                psFromDate = DateTime.Parse(sFromDayPart & " " & sFromMonthPart & " " & sFromYearPart)
            Catch ex As Exception
                ValidDate = False
                sMessage &= "Incorrect 'From' date"
                lblDateError.Text = "Invalid date: " & sMessage & " "
            End Try
            Try
                psToDate = DateTime.Parse(sToDayPart & " " & sToMonthPart & " " & sToYearPart)
            Catch ex As Exception
                ValidDate = False
                sMessage &= "Incorrect 'To' date"
                lblDateError.Text = "Invalid date: " & sMessage
            End Try
        End If
    End Function
    
    Private Sub ExportCSVData()
        Dim response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        response.Clear()
        ' Add the header that specifies the default filename
        ' for the Download/SaveAs dialog
        response.AddHeader(C_HTTP_HEADER_CONTENT, C_HTTP_ATTACHMENT & psFileNameToExport)
    
        ' Specify that the response is a stream that cannot be read _
        ' by the client and must be downloaded
        response.ContentType = C_HTTP_CONTENT_TYPE_OCTET
    
        Dim _exportContent As String = String.Empty
        If (Not _exportData Is Nothing) AndAlso _exportData.Table.Rows.Count > 0 Then
            Dim dv As Dataview = _exportData
            _exportContent = sConvertDataViewToString(dv)
        End If
        If _exportContent.Length <= 0 Then
            _exportContent = C_ERROR_NO_RESULT
        End If
    
        Dim Encoding As New System.Text.UTF8Encoding
        response.AddHeader(C_HTTP_CONTENT_LENGTH, Encoding.GetByteCount(_exportContent).ToString())
        response.BinaryWrite(Encoding.GetBytes(_exportContent))
        response.Charset = ""
    
        ' Stop execution of the current page
        response.End()
    End Sub
    
    Public Function sConvertDataViewToString(ByVal srcDataView As DataView) As String
        Dim ResultBuilder As StringBuilder
        ResultBuilder = New StringBuilder()
        ResultBuilder.Length = 0
    
        Dim aCol As DataColumn
        For Each aCol In srcDataView.Table.Columns
            ResultBuilder.Append(aCol.ColumnName)
            ResultBuilder.Append(",")
        Next
        If ResultBuilder.Length > 1 Then
            ResultBuilder.Length = ResultBuilder.Length - 1
        End If
        ResultBuilder.Append(Environment.NewLine)
    
        Dim aRow As DataRowView 'DataRow
        For Each aRow In srcDataView 'srcDataView.Rows
            For Each aCol In srcDataView.Table.Columns
                ResultBuilder.Append(aRow(Replace(aCol.ColumnName, ",", " ")))
                ResultBuilder.Append(",")
            Next aCol
            ResultBuilder.Length = ResultBuilder.Length - 1
            ResultBuilder.Append(vbNewLine)
        Next aRow
        If Not ResultBuilder Is Nothing Then
            Return ResultBuilder.ToString()
        Else
            Return String.Empty
        End If
    End Function

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

        btnRunCostCentreReport.Enabled = True
        btnRunProductCodeReport.Enabled = True
        btnRunCategoryReport.Enabled = True
        btnRunSubCategoryReport.Enabled = True
    End Sub
    
    Property psFromDate() As String
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
    
    Property psToDate() As String
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
    
    Property pdvReportDataDataView() As DataView
        Get
            Return _exportData
        End Get
    
        Set(ByVal Value As DataView)
            _exportData = Value
        End Set
    End Property
    
    Public Property psFileNameToExport() As String
        Get
            Dim _filename As String = CType(MyBase.ViewState("Filename"), String)
            If _filename Is Nothing Then
                Return "SprintReportData.csv"
            End If
            Return _filename
        End Get
    
        Set(ByVal Value As String)
            If Value Is Nothing Then
                Throw New ArgumentNullException("FileNameToExport", "Provide valid file name for export.")
            End If
            MyBase.ViewState("Filename") = Value
        End Set
    End Property
    
    Property psReportName() As String
        Get
            Dim o As Object = ViewState("ReportName")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("ReportName") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("_IsProductOwner") = Value
        End Set
    End Property
   
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Consignment Finder</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmConsignmentFinder" runat="server">
            <asp:Panel id="pnlReportCriteria" runat="server" Width="100%">
                <table style="width:100%; font-family:Verdana; font-size:xx-small">
                    <tr >
                        <td style="width:10%"></td>
                        <td style="width:25%"></td>
                        <td style="width:10%"></td>
                        <td style="width:25%"></td>
                        <td style="width:30%"></td>
                    </tr>
                    <tr >
                        <td style="white-space:nowrap" colspan="4">
                            <asp:Label ID="lbl001" runat="server" font-size="Small" font-bold="True">Consignment Finder</asp:Label>
                        </td>
                        <td style="white-space:nowrap">
                            <asp:LinkButton ID="lnkbtnDownload" onclick="btn_DownloadCSVFile_Click" runat="server" ForeColor="Blue" CausesValidation="False" Font-Size="XX-Small" Font-Names="Verdana">download to desktop</asp:LinkButton>
                        </td>
                    </tr>
                    <tr>
                        <td align="right" style="white-space: nowrap">
                        </td>
                        <td style="white-space: nowrap">
                        </td>
                        <td align="right">
                        </td>
                        <td>
                        </td>
                        <td>
                        </td>
                    </tr>
                    <tr id="trProductGroups" runat="server">
                        <td colspan="2" style="white-space: nowrap; height: 14px">
                            &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                                Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                                Visible="False">
                            </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Bold="True"
                                Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                        <td align="right" style="height: 14px">
                        </td>
                        <td style="height: 14px">
                        </td>
                        <td align="right" style="height: 14px">
                            <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                                Text="show product groups" Visible="False" /></td>
                    </tr>
                    <tr>
                        <td align="right" style="white-space: nowrap">
                        </td>
                        <td style="white-space: nowrap">
                        </td>
                        <td align="right">
                        </td>
                        <td>
                        </td>
                        <td>
                        </td>
                    </tr>
                    <tr >
                        <td align="right" style="white-space:nowrap">
                            <asp:Label ID="lbl000" runat="server" font-size="XX-Small" Font-Names="Verdana" >From:</asp:Label>
                        </td>
                        <td style="white-space:nowrap">
                            <asp:DropDownList runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlFromDay">
                                <asp:ListItem Value="0">DAY</asp:ListItem>
                                <asp:ListItem Value="1">1</asp:ListItem>
                                <asp:ListItem Value="2">2</asp:ListItem>
                                <asp:ListItem Value="3">3</asp:ListItem>
                                <asp:ListItem Value="4">4</asp:ListItem>
                                <asp:ListItem Value="5">5</asp:ListItem>
                                <asp:ListItem Value="6">6</asp:ListItem>
                                <asp:ListItem Value="7">7</asp:ListItem>
                                <asp:ListItem Value="8">8</asp:ListItem>
                                <asp:ListItem Value="9">9</asp:ListItem>
                                <asp:ListItem Value="10">10</asp:ListItem>
                                <asp:ListItem Value="11">11</asp:ListItem>
                                <asp:ListItem Value="12">12</asp:ListItem>
                                <asp:ListItem Value="13">13</asp:ListItem>
                                <asp:ListItem Value="14">14</asp:ListItem>
                                <asp:ListItem Value="15">15</asp:ListItem>
                                <asp:ListItem Value="16">16</asp:ListItem>
                                <asp:ListItem Value="17">17</asp:ListItem>
                                <asp:ListItem Value="18">18</asp:ListItem>
                                <asp:ListItem Value="19">19</asp:ListItem>
                                <asp:ListItem Value="20">20</asp:ListItem>
                                <asp:ListItem Value="21">21</asp:ListItem>
                                <asp:ListItem Value="22">22</asp:ListItem>
                                <asp:ListItem Value="23">23</asp:ListItem>
                                <asp:ListItem Value="24">24</asp:ListItem>
                                <asp:ListItem Value="25">25</asp:ListItem>
                                <asp:ListItem Value="26">26</asp:ListItem>
                                <asp:ListItem Value="27">27</asp:ListItem>
                                <asp:ListItem Value="28">28</asp:ListItem>
                                <asp:ListItem Value="29">29</asp:ListItem>
                                <asp:ListItem Value="30">30</asp:ListItem>
                                <asp:ListItem Value="31">31</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlFromMonth">
                                <asp:ListItem Value="0" Selected="True">MONTH</asp:ListItem>
                                <asp:ListItem Value="1">JAN</asp:ListItem>
                                <asp:ListItem Value="2">FEB</asp:ListItem>
                                <asp:ListItem Value="3">MAR</asp:ListItem>
                                <asp:ListItem Value="4">APR</asp:ListItem>
                                <asp:ListItem Value="5">MAY</asp:ListItem>
                                <asp:ListItem Value="6">JUN</asp:ListItem>
                                <asp:ListItem Value="7">JUL</asp:ListItem>
                                <asp:ListItem Value="8">AUG</asp:ListItem>
                                <asp:ListItem Value="9">SEP</asp:ListItem>
                                <asp:ListItem Value="10">OCT</asp:ListItem>
                                <asp:ListItem Value="11">NOV</asp:ListItem>
                                <asp:ListItem Value="12">DEC</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList ID="ddlFromYear" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td align="right">
                            <asp:Label ID="lbl002" runat="server" font-size="XX-Small" Font-Names="Verdana">Cost Centre:</asp:Label>
                        </td>
                        <td >
                            <asp:DropDownList id="drop_CostCentre" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td>
                            <asp:Button ID="btnRunCostCentreReport" runat="server" OnClick="btn_Run_CostCentre_Report_Click" Text="generate report" />
                        </td>
                    </tr>
                    <tr >
                        <td align="right" style="white-space:nowrap">
                            <asp:Label ID="lbl003" runat="server" font-size="XX-Small" Font-Names="Verdana" >To:</asp:Label>
                        </td>
                        <td style="white-space:nowrap">
                            <asp:DropDownList runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlToDay">
                                <asp:ListItem Value="0">DAY</asp:ListItem>
                                <asp:ListItem Value="1">1</asp:ListItem>
                                <asp:ListItem Value="2">2</asp:ListItem>
                                <asp:ListItem Value="3">3</asp:ListItem>
                                <asp:ListItem Value="4">4</asp:ListItem>
                                <asp:ListItem Value="5">5</asp:ListItem>
                                <asp:ListItem Value="6">6</asp:ListItem>
                                <asp:ListItem Value="7">7</asp:ListItem>
                                <asp:ListItem Value="8">8</asp:ListItem>
                                <asp:ListItem Value="9">9</asp:ListItem>
                                <asp:ListItem Value="10">10</asp:ListItem>
                                <asp:ListItem Value="11">11</asp:ListItem>
                                <asp:ListItem Value="12">12</asp:ListItem>
                                <asp:ListItem Value="13">13</asp:ListItem>
                                <asp:ListItem Value="14">14</asp:ListItem>
                                <asp:ListItem Value="15">15</asp:ListItem>
                                <asp:ListItem Value="16">16</asp:ListItem>
                                <asp:ListItem Value="17">17</asp:ListItem>
                                <asp:ListItem Value="18">18</asp:ListItem>
                                <asp:ListItem Value="19">19</asp:ListItem>
                                <asp:ListItem Value="20">20</asp:ListItem>
                                <asp:ListItem Value="21">21</asp:ListItem>
                                <asp:ListItem Value="22">22</asp:ListItem>
                                <asp:ListItem Value="23">23</asp:ListItem>
                                <asp:ListItem Value="24">24</asp:ListItem>
                                <asp:ListItem Value="25">25</asp:ListItem>
                                <asp:ListItem Value="26">26</asp:ListItem>
                                <asp:ListItem Value="27">27</asp:ListItem>
                                <asp:ListItem Value="28">28</asp:ListItem>
                                <asp:ListItem Value="29">29</asp:ListItem>
                                <asp:ListItem Value="30">30</asp:ListItem>
                                <asp:ListItem Value="31">31</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana" ID="ddlToMonth">
                                <asp:ListItem Value="0" Selected="True">MONTH</asp:ListItem>
                                <asp:ListItem Value="1">JAN</asp:ListItem>
                                <asp:ListItem Value="2">FEB</asp:ListItem>
                                <asp:ListItem Value="3">MAR</asp:ListItem>
                                <asp:ListItem Value="4">APR</asp:ListItem>
                                <asp:ListItem Value="5">MAY</asp:ListItem>
                                <asp:ListItem Value="6">JUN</asp:ListItem>
                                <asp:ListItem Value="7">JUL</asp:ListItem>
                                <asp:ListItem Value="8">AUG</asp:ListItem>
                                <asp:ListItem Value="9">SEP</asp:ListItem>
                                <asp:ListItem Value="10">OCT</asp:ListItem>
                                <asp:ListItem Value="11">NOV</asp:ListItem>
                                <asp:ListItem Value="12">DEC</asp:ListItem>
                            </asp:DropDownList>
                            <asp:DropDownList ID="ddlToYear" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td align="right" style="white-space:nowrap">
                            <asp:Label ID="lbl004" runat="server" font-size="XX-Small" Font-Names="Verdana">Product Code:</asp:Label>
                        </td>
                        <td >
                            <asp:DropDownList ID="ddlProductCode" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td>
                            <asp:Button ID="btnRunProductCodeReport" OnClick="btn_Run_ProductCode_Report_Click" runat="server" Text="generate report" />
                        </td>
                    </tr>
                    <tr >
                        <td></td>
                        <td></td>
                        <td align="right">
                            <asp:Label ID="lbl005" runat="server" font-size="XX-Small" Font-Names="Verdana">Category:</asp:Label>
                        </td>
                        <td >
                            <asp:DropDownList ID="ddlCategory" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td>
                            <asp:Button ID="btnRunCategoryReport" OnClick="btn_Run_Category_Report_Click" runat="server" Text="generate report" />
                        </td>
                    </tr>
                    <tr >
                        <td></td>
                        <td></td>
                        <td align="right" style="white-space:nowrap">
                            <asp:Label ID="lbl006" runat="server" font-size="XX-Small" Font-Names="Verdana">Sub Category:</asp:Label>
                        </td>
                        <td >
                            <asp:DropDownList ID="ddlSubCategory" runat="server" ForeColor="Navy" Font-Size="XX-Small" Font-Names="Verdana"></asp:DropDownList>
                        </td>
                        <td>
                            <asp:Button ID="btnRunSubCategoryReport" OnClick="btn_Run_SubCategory_Report_Click" runat="server" Text="generate report" />
                        </td>
                    </tr>
                    <tr >
                        <td colspan="5">
                            &nbsp;<asp:Label runat="server" forecolor="Red" id="lblDateError" font-size="XX-Small" ></asp:Label>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <br />
            <asp:Panel id="pnlReportData" runat="server" Width="100%">
                <asp:Label id="lblReportTitle" runat="server" forecolor="Navy" font-size="X-Small" font-names="Verdana" Font-Bold="True"></asp:Label>
                <br />
                <br />
                <asp:DataGrid id="grid_ReportData" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" OnSortCommand="SortReportColumns" AllowSorting="True" AutoGenerateColumns="False" GridLines="None">
                    <AlternatingItemStyle backcolor="White"></AlternatingItemStyle>
                    <ItemStyle backcolor="LightGray"></ItemStyle>
                    <Columns>
                        <asp:BoundColumn DataField="AWB" SortExpression="AWB" HeaderText="Consignment No">
                            <HeaderStyle forecolor="Blue"></HeaderStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Product Code">
                            <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                            <ItemStyle verticalalign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="CreatedOn" SortExpression="CreatedOn" HeaderText="Date" DataFormatString="{0:dd.MM.yy}">
                            <HeaderStyle forecolor="Blue"></HeaderStyle>
                            <ItemStyle verticalalign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                            <HeaderStyle wrap="False" forecolor="Blue"></HeaderStyle>
                            <ItemStyle verticalalign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="Destination">
                            <HeaderStyle forecolor="Blue"></HeaderStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="ItemsOut" SortExpression="ItemsOut" HeaderText="Quantity" DataFormatString="{0:#,##0}">
                            <HeaderStyle horizontalalign="Right" forecolor="Blue"></HeaderStyle>
                            <ItemStyle horizontalalign="Right" verticalalign="Top"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Cost" SortExpression="Cost" HeaderText="Cost (&#163;)" DataFormatString="{0:#,##0.00}">
                            <HeaderStyle horizontalalign="Right"></HeaderStyle>
                            <ItemStyle horizontalalign="Right"></ItemStyle>
                        </asp:BoundColumn>
                    </Columns>
                </asp:DataGrid>
            </asp:Panel>
            &nbsp;
            <asp:Label id="lblError" runat="server" forecolor="Navy" font-size="XX-Small" font-names="Verdana"></asp:Label>
    </form>
</body>
</html>
