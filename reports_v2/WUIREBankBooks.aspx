<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<script runat="server">

    Dim iDefaultHistoryPeriod As Integer = -1    'last 1 month
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim gnTimeout As Int32
    
    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
    
        lblReportGeneratedDateTime.Text = "Report generated: " & Now().ToString("dd-MMM-yy HH:mm")
        If Not IsPostBack Then
            pbIsProductOwner = False
            trProductGroups.Visible = pbProductOwners
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
            BindData()
        End If
    End Sub
    
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
        gnTimeout = Server.ScriptTimeout
        Server.ScriptTimeout = 3600
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
        Server.ScriptTimeout = gnTimeout
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
        Call RunReport()
    End Sub

    Protected Sub RunReport()
        lblFromErrorMessage.Text = ""
        lblToErrorMessage.Text = ""
        
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
            If Not GetDateRange() Then
                Exit Sub
            End If
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            lblReportGeneratedDateTime.Visible = True
            Call BindData()
            pnlData.Visible = True
            pbIsDisplayingData = True

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

    Protected Sub BindData()
        Dim sSQL As String = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  EntryDateTime, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), EntryDateTime, 108)),1,5) 'Date Entered', ConsignmentNumber 'Consignment #', TermID 'Terminal ID', AgentName 'Agent', Address, Town, Region, BookNumber 'Book Number', BookFrom 'From', BookTo 'To' FROM ClientData_WUIRE_BankBookTracking WHERE IsDeleted = 0 AND EntryDateTime >= '" & psFromDate & "' AND EntryDateTime <= '" & psToDate & "' ORDER BY [id]"
        Dim dtEntries As DataTable = ExecuteQueryToDataTable(sSQL)
        
        If dtEntries.Rows.Count > 0 Then
            gvItems.DataSource = dtEntries
            gvItems.DataBind()
            btnExportToExcel1.Visible = True
            btnExportToExcel2.Visible = True
        Else
            btnExportToExcel1.Visible = False
            btnExportToExcel2.Visible = False
            lblError.Text = "No data found for this date range"
            lblReportGeneratedDateTime.Visible = False
        End If
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

    Protected Sub btnExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Call ExportProductDetails()
        Call Export()
    End Sub

    Protected Sub Export()
        Dim sFilename As String = psFromDate.Substring(0, psFromDate.IndexOf(" ")) & " to " & psToDate.Substring(0, psFromDate.IndexOf(" ")) & " "
        Dim sSQL As String = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  EntryDateTime, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), EntryDateTime, 108)),1,5) 'Date Entered', ConsignmentNumber 'Consignment #', TermID 'Terminal ID', AgentName 'Agent', Address, Town, Region, BookNumber 'Book Number', BookFrom 'From', BookTo 'To' FROM ClientData_WUIRE_BankBookTracking WHERE IsDeleted = 0 AND EntryDateTime >= '" & psFromDate & "' AND EntryDateTime <= '" & psToDate & "' ORDER BY [id]"
        Dim dtEntries As DataTable = ExecuteQueryToDataTable(sSQL)
        If dtEntries.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=WU_Ireland_Bank_Books_(" & sFilename & ".csv")
    
            Dim oDataColumn As DataColumn
            Dim sItem As String
    
            Dim IgnoredItems As New ArrayList
    
            For Each oDataColumn In dtEntries.Columns  ' write column header
                If Not IgnoredItems.Contains(oDataColumn.ColumnName) Then
                    Response.Write(oDataColumn.ColumnName)
                    Response.Write(",")
                End If
            Next
            Response.Write(vbCrLf)
    
            For Each dr As DataRow In dtEntries.Rows
                For Each oDataColumn In dtEntries.Columns
                    If Not IgnoredItems.Contains(oDataColumn.ColumnName) Then
                        sItem = (dr(oDataColumn.ColumnName).ToString)
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
        gvItems.PageIndex = e.NewPageIndex
        gvItems.DataBind()
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
        
        gvItems.PageIndex = 0
        
        pnlData.Visible = False
        pbIsDisplayingData = False
        lblError.Text = String.Empty

        ddlProductGroup.Enabled = True
    End Sub

    Protected Function GetDateRange() As Boolean
        GetDateRange = True
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
        psFromDate = psFromDate + " 00:00"
        psToDate = psToDate + " 23:59"
        Dim dtFromDate As Date = Date.Parse(psFromDate)
        Dim dtToDate As Date = Date.Parse(psToDate)
        Dim nDays As Int32 = DateDiff(DateInterval.Day, dtFromDate, dtToDate)
        If nDays < 0 Then
            WebMsgBox.Show("From date appears to be later than To date.")
            GetDateRange = False
        ElseIf nDays > 90 Then
            WebMsgBox.Show("To protect server resouces, the maximum allowed interval between dates is 90 days - to request a report for a longer intervals please contact Transworld.")
            GetDateRange = False
        End If
    End Function
    
    Protected Sub ddlRows_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gvItems.PageIndex = 0
        gvItems.PageSize = ddlRows.SelectedValue
        Call BindData()
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
    <title>Western Union Ireland - Bank Books</title>
    <link rel="stylesheet" type="text/css" href="../Reports.css" />
</head>
<body>
    <form id="frmItemBookingHistoryReport" runat="server">
    <table id="tblDateRangeSelector" runat="server" visible="true" width="100%">
        <tr id="Tr1" runat="server">
            <td colspan="4" style="white-space: nowrap">
                <asp:Label ID="lblPageHeading" runat="server" ForeColor="Silver" Font-Size="Small"
                    Font-Bold="True" Font-Names="Arial">Western Union IRELAND - Bank Books</asp:Label>
            </td>
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
                </asp:DropDownList>
                <asp:Label ID="lblProductGroup" runat="server" Font-Bold="True" Font-Names="Verdana"
                    Font-Size="X-Small"></asp:Label>
            </td>
            <td style="width: 50%; white-space: nowrap">
            </td>
            <td style="width: 25%; white-space: nowrap">
            </td>
            <td style="width: 15%; white-space: nowrap">
                <asp:Button ID="btnShowProductGroups" runat="server" OnClick="btnShowProductGroups_Click"
                    Text="show product groups" Visible="False" />
            </td>
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
            <td style="width: 10%; white-space: nowrap">
                From:
                <asp:TextBox ID="tbFromDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90"
                    runat="server">
                </asp:TextBox>
                <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                    <img alt="" id="Img1" src="../images/SmallCalendar.gif" runat="server" border="0"
                        ie:visible="true" visible="false" /></a> <span id="spnDateExample1" runat="server"
                            visible="true" class="informational light">(eg&nbsp;12-Jan-2007)</span>
            </td>
            <td style="width: 50%; white-space: nowrap">
                <span class="informational dark"></span>&nbsp;&nbsp; <span class="informational dark">
                    To:</span>
                <asp:TextBox ID="tbToDate" Font-Names="Verdana" Font-Size="XX-Small" Width="90" runat="server" />
                <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')">
                    <img alt="" id="Img2" src="../images/SmallCalendar.gif" runat="server" border="0"
                        ie:visible="true" visible="false" />
                </a><span id="spnDateExample2" runat="server" visible="true" class="informational light">
                    (eg&nbsp;12-Jan-2008)</span>
            </td>
            <td style="width: 25%; white-space: nowrap">
                <asp:Button ID="btnRunReport1" runat="server" Text="generate report" Visible="true"
                    OnClick="btnRunReport_Click" Width="170px" />
                <asp:Button ID="btnExportToExcel1" runat="server" OnClick="btnExportToExcel_Click"
                    Text="export to excel" Visible="False" />
                <asp:Button ID="btnReselectReportFilter1" runat="server" Text="re-select report period"
                    Visible="false" OnClick="btnReselectReportFilter_Click" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </td>
            <td style="width: 15%; white-space: nowrap">
                <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                    ToolTip="toggles between calendar style dates and dropdown style dates" />
            </td>
        </tr>
        <tr runat="server" visible="false" id="DropdownInterface">
            <td style="width: 10%; white-space: nowrap">
                From: &nbsp;<asp:DropDownList ID="ddlFromDay" runat="server" Font-Names="Verdana"
                    Font-Size="XX-Small">
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
                </asp:DropDownList>
                &nbsp;<asp:DropDownList ID="ddlFromMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                </asp:DropDownList>
                &nbsp;<asp:DropDownList ID="ddlFromYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                </asp:DropDownList>
            </td>
            <td style="width: 50%; white-space: nowrap">
                <span class="informational dark"></span>&nbsp; &nbsp;&nbsp; <span class="informational dark">
                    To:</span> &nbsp;<asp:DropDownList ID="ddlToDay" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small">
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
                    </asp:DropDownList>
                &nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
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
                </asp:DropDownList>
                &nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small">
                </asp:DropDownList>
                &nbsp;
            </td>
            <td style="width: 25%; white-space: nowrap">
                <asp:Button ID="btnRunReport2" runat="server" Text="generate report" OnClick="btnRunReport_Click"
                    Width="170px" />
                <asp:Button ID="btnExportToExcel2" runat="server" OnClick="btnExportToExcel_Click"
                    Text="export to excel" Visible="False" />
                <asp:Button ID="btnReselectReportFilter2" runat="server" Text="re-select report period"
                    Visible="false" OnClick="btnReselectReportFilter_Click" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            </td>
            <td style="width: 15%; white-space: nowrap">
                <asp:LinkButton ID="lnkbtnToggleSelectionStyle2" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click"
                    ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface" />
            </td>
        </tr>
        <tr runat="server" visible="true" id="DateValidationMessages">
            <td>
                <asp:RegularExpressionValidator ID="revFromDate" runat="server" ControlToValidate="tbFromDate"
                    ErrorMessage=" - invalid format for date - use dd-mmm-yy" Font-Names="Verdana"
                    Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                    SetFocusOnError="True" ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator>
                <asp:RangeValidator ID="rvFromDate" runat="server" ControlToValidate="tbFromDate"
                    CultureInvariantValues="True" ErrorMessage=" - year before 2000, after 2020, or not a valid date!"
                    Font-Names="Verdana" Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1"
                    ValidationGroup="CalendarInterface" EnableClientScript="False" Type="Date"></asp:RangeValidator>
                <asp:Label ID="lblFromErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                    Font-Size="XX-Small" ForeColor="Red"></asp:Label>
            </td>
            <td>
                <asp:RegularExpressionValidator ID="RegularevToDate" runat="server" ControlToValidate="tbToDate"
                    ErrorMessage=" - invalid format for date - use dd-mmm-yy" Font-Names="Verdana"
                    Font-Size="XX-Small" ValidationExpression="^\d\d-(jan|Jan|JAN|feb|Feb|FEB|mar|Mar|MAR|apr|Apr|APR|may|May|MAY|jun|Jun|JUN|jul|Jul|JUL|aug|Aug|AUG|sep|Sep|SEP|oct|Oct|OCT|nov|Nov|NOV|dec|Dec|DEC)-\d(\d+)"
                    ValidationGroup="CalendarInterface"></asp:RegularExpressionValidator><asp:RangeValidator
                        ID="rvToDate" runat="server" ControlToValidate="tbToDate" CultureInvariantValues="True"
                        ErrorMessage=" - year before 2000, after 2020, or not a valid date!" Font-Names="Verdana"
                        Font-Size="XX-Small" MaximumValue="2019/1/1" MinimumValue="2000/1/1" ValidationGroup="CalendarInterface"
                        EnableClientScript="False" Type="Date"></asp:RangeValidator>
                <asp:Label ID="lblToErrorMessage" runat="server" Font-Names="Verdana,Sans-Serif"
                    Font-Size="XX-Small" ForeColor="Red"></asp:Label>
            </td>
            <td>
                &nbsp;
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
    </table>
    <asp:Panel ID="pnlData" runat="server" Visible="false" Width="100%">
        <asp:GridView ID="gvItems" runat="server" Width="100%" OnPageIndexChanging="gvItems_PageIndexChanging"
            AllowPaging="True" CellPadding="2"
            Font-Names="Verdana" Font-Size="XX-Small" >
            <PagerStyle Font-Bold="False" Font-Names="Verdana" Font-Size="Small" HorizontalAlign="Center" />
            <AlternatingRowStyle BackColor="WhiteSmoke" />
            <PagerSettings Position="TopAndBottom" />
        </asp:GridView>
        <br />
        <table>
            <tr>
                <td style="width: 300px">
                    &nbsp;<asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" Font-Size="XX-Small"
                        Font-Names="Verdana" ForeColor="Green" Visible="false"/>
                </td>
                <td style="width: 500px">
                    Items per page:
                    <asp:DropDownList ID="ddlRows" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlRows_SelectedIndexChanged">
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>50</asp:ListItem>
                        <asp:ListItem>100</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <br />
    <asp:Label ID="lblError" runat="server" Font-Names="Arial" Font-Size="XX-Small" ForeColor="red"></asp:Label>
    </form>
</body>
</html>
