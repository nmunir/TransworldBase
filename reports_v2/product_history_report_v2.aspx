<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>
<script runat="server">

    '   Product History Report
    '
    '   Shows Goods In and Goods Out for chosen Product
    '   Apr/May 06 - updated by CN to fix cosmetic problems & use improved movement extraction sproc
    '   NOTE: Cannot get date formatting string to be recognised in product listing
    
    '   TO DO: Skin / CSS all visual items
    
    ' compare VSOE product_history_report_v2.aspx with Copy of product_history_report_v2.aspx
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gdvExportData As DataView
    Private gsMonthNames() As String = {"", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
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
                    ' btnShowAllProducts.Visible = False
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
            lblReportGeneratedDateTime.Visible = False
            Call ShowProductSelection()
        End If
        txtSearchCriteriaAllProducts.Attributes.Add("onkeypress", "return clickButton(event,'" + btnGo.ClientID + "')")
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
            btnShowAllProducts.Enabled = False
            btnGo.Enabled = False
            pnSelectedProductGroup = -1
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlProductList.Visible = False
        pnlMovementList.Visible = False
    End Sub
    
    Protected Sub ShowProductSelection()
        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        Call HideAllPanels()
        pnlProductList.Visible = True
    End Sub
    
    Protected Sub ShowMovementList()
        Call HideAllPanels()
        pnlMovementList.Visible = True
    End Sub
    
    Protected Sub btn_GoToProductList_click(ByVal s As Object, ByVal e As EventArgs)
        dgProducts.EnableViewState = True
        Call ShowProductSelection()
    End Sub
    
    Protected Sub btn_SelectProduct_click(ByVal s As Object, ByVal e As EventArgs)
        dgMovements.EnableViewState = False
        tabChart.Visible = True
        Call ShowProductSelection()
    End Sub
    
    Protected Sub ProductGrid_item_click(ByVal sender As Object, ByVal e As DataGridCommandEventArgs)
        If e.CommandSource.CommandName = "info" Then
            Dim dgi As DataGridItem = e.Item
            lblProductTitle.Text = dgi.Cells(2).Text & " - " & dgi.Cells(4).Text
            
            Dim itemCell As TableCell = e.Item.Cells(0)
            plProductKey = CLng(itemCell.Text)
            BindProductMovementsGrid()
            ShowMovementList()
        End If
    End Sub
    
    Protected Sub imgbtnExportProductDetails_click(ByVal sender As Object, ByVal e As ImageClickEventArgs) '!!!
        Call ExportProductDetails()
    End Sub

    Protected Sub btnExportProductList_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call ExportProductDetails()
    End Sub

    Protected Sub BindProductGrid(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataSet As New DataSet()
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetAllCustProds2", oConn)
        Dim sSearchCriteria As String = Session("ProductSearchCriteria")
        If sSearchCriteria = "" Then
            sSearchCriteria = "_"
        End If
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SearchCriteria", SqlDbType.NVarChar, 50))
        oAdapter.SelectCommand.Parameters("@SearchCriteria").Value = sSearchCriteria

        Try
            oAdapter.Fill(oDataSet, "Movements")
            Dim Source As DataView = oDataSet.Tables("Movements").DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                dgProducts.Visible = True
                dgProducts.DataSource = Source
                dgProducts.DataBind()
            Else
                dgProducts.Visible = False
                lblError.Text = "no data found"
                lblReportGeneratedDateTime.Visible = False
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function GetProductQuantity(ByVal nProductKey As Int32) As Integer
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Product_GetQuantityInStock", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = nProductKey
        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                GetProductQuantity = oDataTable.Rows(0).Item(0)
            Else
                GetProductQuantity = -1
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub BindProductMovementsGrid()
        Const CHART_RANGE As Integer = 60
        Dim sSimpleEncode As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        'Dim oAdapter As New SqlDataAdapter("spASPNET_ProductMovement_History",oConn)
        Dim oAdapter As New SqlDataAdapter("spReportMovementHistory3", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductKey", SqlDbType.Int, 4))
        oAdapter.SelectCommand.Parameters("@ProductKey").Value = plProductKey
        
        Try
            oAdapter.Fill(oDataTable)
            
            If oDataTable.Rows.Count > 0 Then
                Dim nIn As Integer = 0, nOut As Integer = 0
                For Each dr As DataRow In oDataTable.Rows
                    nIn += dr("ItemsIn")
                    nOut += dr("ItemsOut")
                Next
                Dim nBalance As Integer = nIn - nOut
                lblTotalIn.Text = nIn
                lblTotalOut.Text = nOut
                lblBalance.Text = nBalance
                lblInStock.Text = GetProductQuantity(plProductKey)
                
                Try
                    Dim lstItemsOut As New List(Of Integer)
                    Dim nMinValue As Integer = Integer.MaxValue
                    Dim nMaxValue As Integer = Integer.MinValue
                    Dim nDifference As Integer
                    Dim dtStartDate As Date = Date.MinValue
                    Dim dtMiddleDate As Date
                    Dim dtEndDate As Date = Date.MinValue
                    Dim tsTimeSpan As TimeSpan
                    Dim lTimeSpanTicks As Long
                    Dim nItemsOut As Integer

                    For Each dr As DataRow In oDataTable.Rows
                        nItemsOut = dr("ItemsOut")
                        If nItemsOut > 0 Then
                            lstItemsOut.Add(nItemsOut)
                            If nItemsOut < nMinValue Then
                                nMinValue = nItemsOut
                            End If
                            If nItemsOut > nMaxValue Then
                                nMaxValue = nItemsOut
                            End If
                            If dtStartDate = Date.MinValue Then
                                dtStartDate = dr("MovementDate")
                            End If
                            dtEndDate = dr("MovementDate")
                        End If
                    Next
                    tsTimeSpan = dtEndDate - dtStartDate
                    lTimeSpanTicks = tsTimeSpan.Ticks
                    lTimeSpanTicks = lTimeSpanTicks / 2
                    tsTimeSpan = TimeSpan.FromTicks(lTimeSpanTicks)
                    dtMiddleDate = dtStartDate + tsTimeSpan
                
                    If lstItemsOut.Count > 0 Then
                        nDifference = nMaxValue - nMinValue
                        If nDifference = 0 Then
                            nDifference = 1
                        End If
                        Dim bRemoveIdenticalSequences As Boolean = lstItemsOut.Count > 1000
                        Dim nMostRecentValue As Integer = -99
                        Dim lstItemsOutAdjusted As New List(Of Integer)
                        Dim sItemsOutAdjusted As String
                        Dim sSimpleEncoding As String = String.Empty
                    
                        For Each n As Integer In lstItemsOut
                            Dim nAdjustedValue As Integer = (CHART_RANGE * ((n - nMinValue) / nDifference)) + 1
                            If nAdjustedValue <> nMostRecentValue Then
                                nMostRecentValue = nAdjustedValue
                                lstItemsOutAdjusted.Add(nAdjustedValue)
                                sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                            Else
                                If Not bRemoveIdenticalSequences Then
                                    lstItemsOutAdjusted.Add(nAdjustedValue)
                                    sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                                End If
                            End If
                        Next
                    
                        sItemsOutAdjusted = CommaSeparateList(lstItemsOutAdjusted)
                    
                        'If sItemsOutAdjusted.Length > 1900 Then
                        If sSimpleEncoding.Length > 1900 Then
                            lstItemsOutAdjusted.Clear()
                            sSimpleEncoding = String.Empty
                            For Each n As Integer In lstItemsOut
                                Dim nAdjustedValue As Integer = (CHART_RANGE * ((n - nMinValue) / nDifference)) + 1
                                If Not ((nAdjustedValue = nMostRecentValue) Or (nAdjustedValue - 1 = nMostRecentValue) Or (nAdjustedValue + 1 = nMostRecentValue)) Then
                                    nMostRecentValue = nAdjustedValue
                                    lstItemsOutAdjusted.Add(nAdjustedValue)
                                    sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                                Else
                                    If Not bRemoveIdenticalSequences Then
                                        lstItemsOutAdjusted.Add(nAdjustedValue)
                                        sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                                    End If
                                End If
                            Next
                        End If
                    
                        sItemsOutAdjusted = CommaSeparateList(lstItemsOutAdjusted)
                    
                        'If sItemsOutAdjusted.Length > 1900 Then
                        If sSimpleEncoding.Length > 1900 Then
                            lstItemsOutAdjusted.Clear()
                            sSimpleEncoding = String.Empty
                            For Each n As Integer In lstItemsOut
                                Dim nAdjustedValue As Integer = (CHART_RANGE * ((n - nMinValue) / nDifference)) + 1
                                If Not ((nAdjustedValue = nMostRecentValue) Or (nAdjustedValue - 1 = nMostRecentValue) Or (nAdjustedValue - 2 = nMostRecentValue) Or (nAdjustedValue + 1 = nMostRecentValue) Or (nAdjustedValue + 2 = nMostRecentValue)) Then
                                    nMostRecentValue = nAdjustedValue
                                    lstItemsOutAdjusted.Add(nAdjustedValue)
                                    sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                                Else
                                    If Not bRemoveIdenticalSequences Then
                                        lstItemsOutAdjusted.Add(nAdjustedValue)
                                        sSimpleEncoding += sSimpleEncode(nAdjustedValue)
                                    End If
                                End If
                            Next
                        End If

                        sItemsOutAdjusted = CommaSeparateList(lstItemsOutAdjusted)

                        Dim sbChart As New StringBuilder
                        Dim bIsBarChart As Boolean = False
                        sbChart.Append("http://chart.apis.google.com/chart?")
                        If lstItemsOut.Count <= 16 Then
                            sbChart.Append("cht=bvs")
                            bIsBarChart = True
                        Else
                            sbChart.Append("cht=lc")
                        End If
                        sbChart.Append("&")
                        sbChart.Append("chs=800x250")
                        sbChart.Append("&")
                        sbChart.Append("chxt=x,y")
                        sbChart.Append("&")
                        sbChart.Append("chd=")
                        'sbChart.Append("t:")
                        sbChart.Append("s:")
                        'sbChart.Append(CommaSeparateList(lstItemsOutAdjusted))
                        sbChart.Append(sSimpleEncoding)
                        sbChart.Append("&")
                        sbChart.Append("chco=008000")
                        sbChart.Append("&")
                        sbChart.Append("chm=B,A5CE84,0,0,0")
                        sbChart.Append("&")
                        sbChart.Append("chxl=")
                        sbChart.Append("0:")
                        sbChart.Append("|")
                        If bIsBarChart Then
                            sbChart.Append(DisplayDate(dtStartDate))
                            sbChart.Append("|".PadRight(Int((lstItemsOutAdjusted.Count - 2) / 2), "|"))
                            If Not (DisplayDate(dtMiddleDate) = DisplayDate(dtStartDate) Or DisplayDate(dtMiddleDate) = DisplayDate(dtEndDate)) Then
                                sbChart.Append(DisplayDate(dtMiddleDate))
                            End If
                            sbChart.Append("|".PadRight(Int((lstItemsOutAdjusted.Count - 2) / 2), "|"))
                            sbChart.Append(DisplayDate(dtEndDate))
                            sbChart.Append("|")
                        Else
                            sbChart.Append(DisplayDate(dtStartDate))
                            sbChart.Append("|")
                            If Not (DisplayDate(dtMiddleDate) = DisplayDate(dtStartDate) Or DisplayDate(dtMiddleDate) = DisplayDate(dtEndDate)) Then
                                sbChart.Append(DisplayDate(dtMiddleDate))
                                sbChart.Append("|")
                            End If
                            sbChart.Append(DisplayDate(dtEndDate))
                            sbChart.Append("|")
                        End If
                        sbChart.Append("1:")
                        sbChart.Append("|")
                        If lstItemsOutAdjusted.Count >= 3 Then
                            sbChart.Append(nMinValue.ToString)
                        End If
                        Dim nIntermediateValue As Integer = Int(nMinValue + ((nMaxValue - nMinValue) / 2))
                        If Not nIntermediateValue = nMinValue Then
                            sbChart.Append("|")
                            sbChart.Append(Int(nMinValue + ((nMaxValue - nMinValue) / 2)).ToString)
                        End If
                        If Not nMaxValue = nMinValue Then
                            sbChart.Append("|")
                            sbChart.Append(nMaxValue.ToString)
                        End If
                        If sbChart.ToString.Length > 2048 Then
                            tabChart.Visible = False
                            lblError.Text = "Too many data points to chart"
                        Else
                            tabChart.Visible = True
                            imgChart.ImageUrl = sbChart.ToString
                        End If
                    Else
                        tabChart.Visible = False
                    End If
                Catch
                    tabChart.Visible = False
                    lblError.Text = "Could not generate chart"
                End Try
                
                dgMovements.Visible = True
                dgMovements.DataSource = oDataTable
                dgMovements.DataBind()
            Else
                dgMovements.Visible = False
                tabChart.Visible = False
                lblTotalIn.Text = 0
                lblTotalOut.Text = 0
                lblBalance.Text = 0
                lblInStock.Text = 0
                lblError.Text = "... no data found"
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Private Function DisplayDate(ByVal dtDate As Date) As String
        Dim nMonth As Integer = dtDate.Month
        Dim nYear As Integer = dtDate.Year
        DisplayDate = gsMonthNames(nMonth) & " " & nYear.ToString.Substring(2, 2)
    End Function
    
    Private Function CommaSeparateList(ByVal lstList As List(Of Integer)) As String
        Dim s As String = String.Empty
        For i As Integer = 0 To lstList.Count - 2
            s = s & lstList(i) & ","
        Next
        s = s & lstList(lstList.Count - 1)
        CommaSeparateList = s
    End Function
            
    Private Function CommaSeparateList(ByVal lstList As List(Of String)) As String
        Dim s As String = String.Empty
        For i As Integer = 0 To lstList.Count - 2
            s = s & lstList(i) & ","
        Next
        s = s & lstList(lstList.Count - 1)
        CommaSeparateList = s
    End Function

    Protected Sub SortProductColumns(ByVal Source As Object, ByVal E As DataGridSortCommandEventArgs)
        BindProductGrid(E.SortExpression)
    End Sub
    
    Public Function ConvertDataViewToCSVString(ByVal oDataView As DataView) As String
        Dim ResultBuilder As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRowView

        For Each oDataColumn In oDataView.Table.Columns         ' column headings in line 1
            ResultBuilder.Append(oDataColumn.ColumnName)
            ResultBuilder.Append(",")
        Next
        If ResultBuilder.Length > 1 Then
            ResultBuilder.Length = ResultBuilder.Length - 1
        End If
        ResultBuilder.Append(Environment.NewLine)
    
        For Each oDataRow In oDataView
            For Each oDataColumn In oDataView.Table.Columns
                ResultBuilder.Append(oDataRow(Replace(oDataColumn.ColumnName, ",", " ")))  ' replace any commas with spaces
                ResultBuilder.Append(",")
            Next oDataColumn
            ResultBuilder.Length = ResultBuilder.Length - 1
            ResultBuilder.Append(Environment.NewLine)
        Next oDataRow

        If Not ResultBuilder Is Nothing Then
            Return ResultBuilder.ToString()
        Else
            Return String.Empty
        End If
    End Function

    Protected Sub ExportProductDetails()
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_Customer_ExportProducts2Undeleted", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        lblError.Text = ""

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        Try
            oAdapter.Fill(oDataTable)
            If oDataTable.Rows.Count > 0 Then
                Response.Clear()
                'Response.ContentType = "Application/x-msexcel"
                Response.ContentType = "text/csv"
                Response.AddHeader("Content-Disposition", "attachment; filename=product_list.csv")
                Dim IgnoredItems As New ArrayList
                IgnoredItems.Add("UserKey")
                IgnoredItems.Add("CurrentEncryptedPassword")
    
                For Each dc As DataColumn In oDataTable.Columns
                    If Not IgnoredItems.Contains(dc.ColumnName) Then
                        Response.Write(dc.ColumnName)
                        Response.Write(",")
                    End If
                Next
                Response.Write(vbCrLf)
    
                For Each dr As DataRow In oDataTable.Rows
                    For Each dc As DataColumn In oDataTable.Columns
                        If Not IgnoredItems.Contains(dc.ColumnName) Then
                            Dim sItem As String = (dr(dc.ColumnName).ToString)
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
                dgProducts.Visible = False
                lblError.Text = "no data found"
            End If
        Catch ex As SqlException
            lblError.Text = ex.ToString
        Finally
            oConn.Close()
        End Try
    End Sub

    Protected Sub btnShowAllProducts_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        dgMovements.EnableViewState = True
        txtSearchCriteriaAllProducts.Text = ""
        Session("ProductSearchCriteria") = txtSearchCriteriaAllProducts.Text
        BindProductGrid("ProductCode")
        lblReportGeneratedDateTime.visible = True
    End Sub

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ProductSearchCriteria") = txtSearchCriteriaAllProducts.Text
        BindProductGrid("ProductCode")
        lblReportGeneratedDateTime.visible = True
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
        btnShowAllProducts.Enabled = True
        btnGo.Enabled = True
    End Sub
    
    Property plProductKey() As Long
        Get
            Dim o As Object = ViewState("PHR_ProductKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CLng(o)
        End Get
        Set(ByVal Value As Long)
            ViewState("PHR_ProductKey") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("PHR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("PHR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("PHR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PHR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("PHR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("PHR_IsProductOwner") = Value
        End Set
    End Property
   
    Protected Sub lnkbtnHideChart_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lb As LinkButton = sender
        tabChart.Visible = False
    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Product History Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="Form1" runat="server">
        <asp:Panel id="pnlProductList" runat="server" Width="100%">
            <table width="100%">
                <tr>
                    <td valign="Bottom" width="5%" style="height: 45px"></td>
                    <td Wrap="False" width="50%" style="height: 45px">
                        <asp:Label ID="Label1" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Verdana">Product History Report</asp:Label>
                        <br /><br />
                    </td>
                    <td Wrap="False" align="Right" width="45%" style="height: 45px"></td>
                </tr>
                <tr>
                    <td style="height: 14px">
                    </td>
                    <td style="height: 14px" wrap="False">
                    </td>
                    <td style="height: 14px">
                    </td>
                </tr>
                <tr id="trProductGroups" runat="server">
                    <td>
                    </td>
                    <td wrap="False">
                        <asp:DropDownList ID="ddlProductGroup" runat="server" AutoPostBack="True" Font-Names="Verdana"
                            Font-Size="XX-Small" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged"
                            Visible="False">
                        </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Bold="True"
                            Font-Names="Verdana" Font-Size="X-Small"></asp:Label></td>
                    <td align="right">
                        <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                            Text="show product groups" Visible="False" /></td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td wrap="False">
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td Wrap="False" colspan="2">
                <asp:Button ID="btnShowAllProducts"
                     runat="server"
                     Text="show all products"
                      OnClick="btnShowAllProducts_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:Label ID="Label2" runat="server" forecolor="Gray" font-size="XX-Small" font-names="Verdana">Search:</asp:Label> <asp:TextBox runat="server" Height="20px" Width="100px" Font-Size="XX-Small" Font-Names="Verdana" ID="txtSearchCriteriaAllProducts"></asp:TextBox>
                        &nbsp;
                        <asp:Button ID="btnGo" runat="server" Text="go" OnClick="btnGo_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnExportProductList" runat="server" Text="export product list to excel"
                                    OnClick="btnExportProductList_Click" />
                    </td>
                </tr>
            </table>
            <br />
            <asp:DataGrid id="dgProducts" runat="server" Width="100%" Font-Names="Arial" Font-Size="XX-Small" OnItemCommand="ProductGrid_item_click" CellSpacing="4" AutoGenerateColumns="False" GridLines="None" AllowSorting="True" OnSortCommand="SortProductColumns">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" forecolor="Blue"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticProductKey" SortExpression="LogisticProductKey" HeaderText="No." DataFormatString="{0:000000}">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn>
                        <ItemTemplate>
                            <asp:ImageButton id="ImageButton1" runat="server" CommandName="info" ToolTip="show history for this product" ImageUrl="../images/icon_arrow.gif"></asp:ImageButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="ProductCode" SortExpression="ProductCode" HeaderText="Code">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDate" SortExpression="ProductDate" HeaderText="Date" DataFormatString="{0:dd-MMM-yy}"></asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDescription" SortExpression="ProductDescription" HeaderText="Description"></asp:BoundColumn>
                    <asp:BoundColumn DataField="ProductDepartmentId" SortExpression="ProductDepartmentId" HeaderText="Department">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="LanguageId" SortExpression="LanguageId" HeaderText="Language">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ArchiveFlag" SortExpression="ArchiveFlag" HeaderText="Archive Flag">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
                <AlternatingItemStyle BackColor="#DDDDDD" />
            </asp:DataGrid>
            <br />
            &nbsp;<asp:Label ID="lblReportGeneratedDateTime" runat="server" font-size="XX-Small" font-names="Verdana,Sans-Serif" forecolor="Green"></asp:Label>
        </asp:Panel>
        <asp:Panel id="pnlMovementList" runat="server" Width="100%">
            <asp:Table id="Table5" runat="server" width="100%">
                <asp:TableRow>
                    <asp:TableCell VerticalAlign="Bottom" width="5%">
                        &nbsp;&nbsp;<asp:Image ID="Image1" runat="server" ImageUrl="../images/icon_back.gif" Visible="false"></asp:Image>
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" width="50%">
                        <asp:Label ID="lblProductTitlePreamble" runat="server" Text="History for product: " forecolor="#0000C0" font-size="X-Small" font-bold="True" font-names="Arial"></asp:Label><asp:Label ID="lblProductTitle" runat="server" forecolor="#0000C0" font-size="X-Small" font-bold="True" font-names="Arial"></asp:Label>

                        <asp:LinkButton ID="LinkButton1" runat="server" visible="false" ForeColor="Blue" Font-Size="X-Small" Font-Names="Arial" onclick="btn_SelectProduct_click">re-select product</asp:LinkButton>
                    </asp:TableCell>
                    <asp:TableCell Wrap="False" HorizontalAlign="Right" width="45%">
                                    <asp:Button ID="btnSelectAnotherProduct"
                                                runat="server"
                                                Text="select another product"
                                                OnClick="btn_SelectProduct_Click" />
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <table id="tabChart" runat="server" style="width: 100%">
                <tr>
                    <td style="width:90%">
                        <asp:Image ID="imgChart" runat="server" />
                    </td>
                    <td style="width:10%">
                        &nbsp;
                        <asp:LinkButton ID="lnkbtnHideChart" runat="server" 
                            onclick="lnkbtnHideChart_Click">hide&nbsp;chart</asp:LinkButton>
                    </td>
                </tr>
            </table>
            <br />
            <asp:DataGrid id="dgMovements" runat="server" Width="100%" Font-Names="Arial" Font-Size="XX-Small" CellSpacing="7" AutoGenerateColumns="False" GridLines="None">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" forecolor="Blue"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticBookingKey" SortExpression="LogisticBookingKey" HeaderText="LogisticBookingKey">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="MovementDate" SortExpression="MovementDate" HeaderText="Date" DataFormatString="{0:dd-MMM-yy}">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ItemsIn" SortExpression="ItemsIn" HeaderText="In" DataFormatString="{0:#,##0}">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ItemsOut" SortExpression="ItemsOut" HeaderText="Out" DataFormatString="{0:#,##0}">
                        <HeaderStyle horizontalalign="Right"></HeaderStyle>
                        <ItemStyle horizontalalign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn></asp:TemplateColumn>
                    <asp:BoundColumn DataField="CneeName" SortExpression="CneeName" HeaderText="Company">
                        <ItemStyle wrap="True"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeAddr1" SortExpression="CneeAddr1" HeaderText="Addr 1">
                        <ItemStyle wrap="True"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="City"></asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country"></asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
                <AlternatingItemStyle BackColor="#DDDDDD" />
            </asp:DataGrid><br />
            <asp:Label ID="lblLegendTotalIn" runat="server" Text="Total In:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:Label ID="lblTotalIn" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            &nbsp;&nbsp; &nbsp;<asp:Label ID="lblLegendTotalOut" runat="server" Text="Total Out:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:Label ID="lblTotalOut" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            &nbsp; &nbsp;
            <asp:Label ID="lblLegendBalance" runat="server" Text="Balance:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:Label ID="lblBalance" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            &nbsp; &nbsp;
            <asp:Label ID="lblLegendInStock" runat="server" Text="In Stock:" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
            <asp:Label ID="lblInStock" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label><br />
        </asp:Panel>
        &nbsp;
        <asp:Label id="lblError" runat="server" font-size="XX-Small" font-names="Arial" forecolor="red"/>
    </form>
</body>
</html>