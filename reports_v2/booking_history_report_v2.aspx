<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<script runat="server">

    '   Booking History Report

    Const STYLENAME_CALENDAR As String = "calendar style dates"
    Const STYLENAME_DROPDOWN As String = "dropdown style dates"

    Dim iDefaultHistoryPeriod As Integer = -3    'last 3 months
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Private gdictCountryCodes As Dictionary(Of String, Integer)
    Private gnMinValue As Integer
    Private gnMaxValue As Integer
    Private gnDifference As Integer
    
    Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
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
        
        pnlData.Visible = False
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
         Or _
          (DropdownInterface.Visible _
          And IsDate(ddlFromDay.SelectedItem.Text & ddlFromMonth.SelectedItem.Text & ddlFromYear.SelectedItem.Text) _
          And IsDate(ddlToDay.SelectedItem.Text & ddlToMonth.SelectedItem.Text & ddlToYear.SelectedItem.Text)) Then

            Call GetDateRange()
            spnDateExample1.Visible = False
            spnDateExample2.Visible = False
            imgCalendarButton1.Visible = False
            imgCalendarButton2.Visible = False
            
            lblReportGeneratedDateTime.Visible = True
            
            BindBookingGrid("BookedBy")
            pnlData.Visible = True
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
    
    Private Function GetAdjustedDataPoint(ByVal nDataPoint As Integer) As Integer
        GetAdjustedDataPoint = 99 * ((nDataPoint - gnMinValue) / gnDifference)
    End Function

    Protected Sub BindBookingGrid(ByVal SortField As String)
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataTable As New DataTable
        Dim oAdapter As New SqlDataAdapter("spASPNET_LogisticBooking_AllUserHistory3", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
    
        lblError.Text = ""
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@CustomerKey").Value = Session("CustomerKey")
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ProductGroup", SqlDbType.Int))
        oAdapter.SelectCommand.Parameters("@ProductGroup").Value = pnSelectedProductGroup

        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@FromDate").Value = CDate(psFromDate)
    
        oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.DateTime))
        oAdapter.SelectCommand.Parameters("@ToDate").Value = DateAdd("D", 1, CDate(psToDate))
    
        Try
            'gdictCountryCodes As Dictionary(Of String, Integer)
            oAdapter.Fill(oDataTable)

            gdictCountryCodes = New Dictionary(Of String, Integer)
            Dim sCountryCode As String
            For Each dr As DataRow In oDataTable.Rows
                If Not IsDBNull(dr("CountryId2")) Then
                    sCountryCode = dr("CountryId2")
                    If gdictCountryCodes.ContainsKey(sCountryCode) Then
                        gdictCountryCodes(sCountryCode) = gdictCountryCodes(sCountryCode) + 1
                    Else
                        gdictCountryCodes.Add(sCountryCode, 1)
                    End If
                End If
            Next
            Dim sbChartData As New StringBuilder
            ' sea is EAF7FE
            sbChartData.Append("http://chart.apis.google.com/chart?cht=t&chtm=world&chf=bg,s,0&chs=440x220&chco=FFFFFF,00FF00,FFFF00,FF0000&chld=")
            
            gnMinValue = Integer.MaxValue
            gnMaxValue = Integer.MinValue
            Dim sCountries As String = String.Empty
            Dim sValues As String = String.Empty
            For Each kv As KeyValuePair(Of String, Integer) In gdictCountryCodes
                'sbChartData.Append(kv.Key)
                If kv.Value < gnMinValue Then
                    gnMinValue = kv.Value
                End If
                If kv.Value > gnMaxValue Then
                    gnMaxValue = kv.Value
                End If
                sCountries = sCountries & kv.Key
                'sValues = sValues & (GetAdjustedDataPoint(kv.Value) + 1) & ","
                'sValues = sValues & kv.Value & ","
            Next
            gnDifference = gnMaxValue - gnMinValue
            If gnDifference = 0 Then
                gnDifference = 1
            End If
            

            For Each kv As KeyValuePair(Of String, Integer) In gdictCountryCodes
                'sCountries = sCountries & kv.Key
                sValues = sValues & (GetAdjustedDataPoint(kv.Value) + 1) & ","
                'sValues = sValues & kv.Value & ","
            Next
            sbChartData.Append(sCountries)
            sValues = sValues.Substring(0, sValues.Length - 1)
            sbChartData.Append("&chd=t:" & sValues)
            If gdictCountryCodes.Count > 0 Then
                imgDistributionMap.ImageUrl = sbChartData.ToString
                imgDistributionMap.Visible = True
            Else
                imgDistributionMap.Visible = False
            End If
            Dim Source As DataView = oDataTable.DefaultView
            Source.Sort = SortField
            If Source.Count > 0 Then
                dgrdBookings.DataSource = Source
                dgrdBookings.DataBind()
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
        BindBookingGrid(E.SortExpression)
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

    Protected Sub btnShowProductGroups_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ShowProductGroups()
        Call ReselectReportFilter()
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
        Call ReselectReportFilter()
    End Sub
    
    Property psToDate() As String
        Get
            Dim o As Object = ViewState("BHR_ToDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("BHR_ToDate") = Value
        End Set
    End Property
    
    Property psFromDate() As String
        Get
            Dim o As Object = ViewState("BHR_FromDate")
            If o Is Nothing Then
                Return ""
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("BHR_FromDate") = Value
        End Set
    End Property
    
    Property pnSelectedProductGroup() As Integer
        Get
            Dim o As Object = ViewState("BHR_SelectedProductGroup")
            If o Is Nothing Then
                Return 2
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("BHR_SelectedProductGroup") = Value
        End Set
    End Property
   
    Property pbProductOwners() As Boolean
        Get
            Dim o As Object = ViewState("BHR_ProductOwners")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_ProductOwners") = Value
        End Set
    End Property
   
    Property pbIsProductOwner() As Boolean
        Get
            Dim o As Object = ViewState("BHR_IsProductOwner")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("BHR_IsProductOwner") = Value
        End Set
    End Property
   
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Booking History Report</title>
    <link rel="stylesheet" type="text/css" href="../css/sprint.css" />
</head>
<body>
    <form id="frmBookingHistoryReport" runat="server">
        <table id="tblDateRangeSelector" runat="server" visible="true">
            <tr>
                <td colspan="4" style="white-space:nowrap; height: 21px;">
                    &nbsp;<asp:Label ID="lblPageHeading" runat="server" forecolor="silver" font-size="Small" font-bold="True" font-names="Arial"> Booking History Report</asp:Label></td>
            </tr>
            <tr>
                <td runat="server" colspan="2" style="white-space: nowrap">
                </td>
                <td style="width: 253px; height: 14px">
                </td>
                <td align="right" style="width: 169px; height: 14px">
                </td>
            </tr>
            <tr runat="server" visible="true" id="trProductGroups">
                <td colspan="2" style="white-space: nowrap" id="tdProductGroups" runat="server">
                    &nbsp;<asp:DropDownList ID="ddlProductGroup" runat="server" Font-Names="Verdana"
                        Font-Size="XX-Small" Visible="False" OnSelectedIndexChanged="ddlProductGroup_SelectedIndexChanged" AutoPostBack="True">
                    </asp:DropDownList><asp:Label ID="lblProductGroup" runat="server" Font-Names="Verdana" Font-Size="X-Small" Font-Bold="True"></asp:Label></td>
                <td style="width: 253px; height: 14px">
                </td>
                <td align="right" style="width: 169px; height: 14px">
                    <asp:Button ID="btnShowProductGroups" runat="server" Text="show product groups" Visible="False" OnClick="btnShowProductGroups_Click" /></td>
            </tr>
            <tr runat="server" visible="true">
                <td style="width: 265px; white-space: nowrap; height: 14px">
                </td>
                <td style="white-space: nowrap; height: 14px">
                </td>
                <td style="width: 253px; height: 14px">
                </td>
                <td style="width: 169px; height: 14px">
                </td>
            </tr>
            <tr runat="server" visible="true" id="CalendarInterface">
                <td style="width: 265px; white-space:nowrap">
                    <span class="informational dark">&nbsp;From:</span>
                        <asp:TextBox ID="tbFromDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                         <a id="imgCalendarButton1" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')">
                            <img id="Img1" src="../images/SmallCalendar.gif" runat="server" border="0" alt="" IE:visible="true" visible="false"
                               /></a><span id="spnDateExample1" runat="server" visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2007)</span>
                </td>
                <td style="white-space:nowrap" >
                    <span class="informational dark">To:</span>
                        <asp:TextBox ID="tbToDate" font-names="Verdana" font-size="XX-Small" Width="90" runat="server"/>
                           <a id="imgCalendarButton2" runat="server" visible="true" href="javascript:;"
                            onclick="window.open('../PopupCalendar4.aspx?textbox=tbToDate','cal','width=300,height=305,left=270,top=180')"><img id="Img2"
                                 src="../images/SmallCalendar.gif" runat="server" border="0" alt="" IE:visible="true" visible="false"
                               /></a>
                    <span id="spnDateExample2" runat="server" visible="true" class="informational light" style="white-space: nowrap">(eg&nbsp;12-Jan-2008)</span>
                </td>
                <td style="width: 253px">
                    <asp:Button ID="btnRunReport1" runat="server" Text="generate report" Visible="true" OnClick="btnRunReport_Click" Width="170px" />
                    <asp:Button ID="btnReselectReportFilter1" runat="server" Text="re-select report filter" Visible="false" OnClick="btnReselectReportFilter_Click" />
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                </td>
                <td style="width: 169px">
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle1" runat="server" OnClick="lnkbtnToggleSelectionStyle_Click" ToolTip="toggles between calendar interface and dropdown interface"></asp:LinkButton>
                </td>
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
                    <asp:LinkButton ID="lnkbtnToggleSelectionStyle2"
                                      runat="server"
                                      OnClick="lnkbtnToggleSelectionStyle_Click"
                                      ToolTip="toggles between easy-to-use calendar interface and clunky dropdown interface"></asp:LinkButton></td>
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
    <asp:Image ID="imgDistributionMap" runat="server" Visible="False" />
    <asp:Panel id="pnlData" runat="server" Visible="false" Width="100%">
            <asp:DataGrid id="dgrdBookings" runat="server" Width="100%" Font-Names="Arial" Font-Size="XX-Small" CellSpacing="7" AutoGenerateColumns="False" GridLines="None" AllowSorting="True" OnSortCommand="SortProductColumns">
                <HeaderStyle font-size="XX-Small" font-names="Verdana" forecolor="Blue"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn Visible="False" DataField="LogisticBookingKey" SortExpression="LogisticBookingKey" HeaderText="No." DataFormatString="{0:000000}">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedOn" SortExpression="BookedOn" HeaderText="Date"  DataFormatString="{0:dd-MMM-yy HH:mm}">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="AWB" SortExpression="AWB" HeaderText="AWB"></asp:BoundColumn>
                    <asp:BoundColumn DataField="BookedBy" SortExpression="BookedBy" HeaderText="Booked By">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CustomerBookingReference" SortExpression="CustomerBookingReference" HeaderText="Reference">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CashOnDelAmount" SortExpression="CashOnDelAmount" HeaderText="Cost">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeName" SortExpression="CneeName" HeaderText="Destination">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CneeTown" SortExpression="CneeTown" HeaderText="City"></asp:BoundColumn>
                    <asp:BoundColumn DataField="CountryName" SortExpression="CountryName" HeaderText="Country">
                        <ItemStyle wrap="False"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
            <br />
            <asp:Label ID="lblReportGeneratedDateTime" runat="server" Text="" font-size="XX-Small" font-names="Verdana, Sans-Serif" forecolor="Green" Visible="false"></asp:Label>
        </asp:Panel>
        <br />
        <asp:Label id="lblError" runat="server" Font-Names="Arial" Font-Size="XX-Small" ForeColor="red"></asp:Label>
    </form>
</body>
</html>
