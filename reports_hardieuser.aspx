<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script runat="server">

    Dim gsConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Dim gColour1 As Drawing.Color = Drawing.Color.White
    Dim gColour2 As Drawing.Color = Drawing.Color.Wheat
    Dim gCurrentColour As Drawing.Color = gColour1
    Dim gsLastAWB As String

    Const CUSTOMER_HARDIE As Integer = 837
    Const CUSTOMER_HARDFR As Int32 = 849

    Protected Sub gvOrders_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim sAWB As String = gvr.Cells(1).Text
            If sAWB <> gsLastAWB = gvr.Cells(1).Text Then
                Call SwapColour()
            End If
            gvr.BackColor = gCurrentColour
            gsLastAWB = sAWB
        End If
    End Sub
    
    Protected Sub SwapColour()
        If gCurrentColour = gColour1 Then
            gCurrentColour = gColour2
        Else
            gCurrentColour = gColour1
        End If
    End Sub
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call SetTitle()
        End If
    End Sub
    
    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "James Hardie Orders Report"
    End Sub
    
    Protected Function IsHardie() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHardie = IIf(gbSiteTypeDefined, gsSiteType = "hardie", nCustomerKey = CUSTOMER_HARDIE)
    End Function
    
    Protected Function IsHardFR() As Boolean
        Dim nCustomerKey As Integer = CInt(Session("CustomerKey"))
        IsHardFR = IIf(gbSiteTypeDefined, gsSiteType = "hardfr", nCustomerKey = CUSTOMER_HARDFR)
    End Function
    
    Protected Function RetrieveOrders(sFromDate As String, sToDate As String) As DataTable
        Dim sSQL As String = "SELECT CAST(REPLACE(CONVERT(VARCHAR(11),  c.CreatedOn, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), c.CreatedOn, 108)),1,5) 'Order Date', AWB 'Consignment #', up.FirstName + ' ' + up.LastName + ' (' + UserID + ')' 'Ordered By', c.CneeCtcName 'Contact Name', c.CneeName 'Company', c.CneeAddr1 'Addr 1', C.CneeTown 'Town', c.CneePostCode 'Post Code', lp.ProductCode 'Product', lp.ProductDescription 'Description', ISNULL(lp.ProductCategory, '') + ' ' + ISNULL(lp.SubCategory, '') + ' ' + ISNULL(lp.SubCategory2, '') 'Category', lm.ItemsOut 'Qty', CAST(REPLACE(CONVERT(VARCHAR(11),  c.WarehouseCutOffTime, 106), ' ', '-') AS varchar(20)) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), c.WarehouseCutOffTime, 108)),1,5) 'Despatched',  ISNULL(PODName, '') + ' ' + ISNULL(PODDate, '') + ' ' + ISNULL(PODTime, '') 'POD' FROM Consignment c INNER JOIN UserProfile up ON c.UserKey = up.[key] INNER JOIN LogisticMovement lm ON c.[Key] = lm.ConsignmentKey INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey WHERE c.CustomerKey IN (837, 849) AND c.StateId = 'WITH_OPERATIONS' AND c.CreatedOn >= '" & sFromDate & "' and c.CreatedOn <= '" & sToDate & "' ORDER BY c.[key]"
        RetrieveOrders = ExecuteQueryToDataTable(sSQL)
    End Function
   
    Protected Sub ShowOrders(sFromDate As String, sToDate As String)
        Dim dtOrders As DataTable = RetrieveOrders(sFromDate, sToDate)
        gvOrders.DataSource = dtOrders
        gvOrders.DataBind()
    End Sub

    Protected Sub ExportOrders(sFromDate As String, sToDate As String)
        Dim dtOrders As DataTable = RetrieveOrders(sFromDate, sToDate)
        If dtOrders.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"
            Response.AddHeader("Content-Disposition", "attachment; filename=hardie_orders_" & DateTime.Now.ToString("dd-MMM-yyyyhhmmss") & ".csv")
    
            Dim oDataColumn As DataColumn
            Dim sItem As String
    
            For Each oDataColumn In dtOrders.Columns  ' write column header
                Response.Write(oDataColumn.ColumnName)
                Response.Write(",")
            Next
            Response.Write(vbCrLf)
    
            For Each dr As DataRow In dtOrders.Rows
                For Each oDataColumn In dtOrders.Columns
                    sItem = (dr(oDataColumn.ColumnName).ToString)
                    sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        End If
    End Sub

    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        Dim sToDate As String = ddlToDay.Text + "-" + ddlToMonth.Text + "-" + ddlToYear.Text + " 23:29"
        Try
            Dim dtFromDate As Date = DateAdd(DateInterval.Day, -30, Date.Parse(sToDate))
            Dim sFromDate As String = dtFromDate.ToString("dd-MMM-yyyy 00:01")
            Call ShowOrders(sFromDate, sToDate)
        Catch ex As Exception
            WebMsgBox.Show("Not a valid date!")
        End Try
    End Sub
    
    Protected Sub btnExport_Click(sender As Object, e As System.EventArgs)
        Dim sToDate As String = ddlToDay.Text + "-" + ddlToMonth.Text + "-" + ddlToYear.Text + " 23:29"
        Try
            Dim dtFromDate As Date = DateAdd(DateInterval.Day, -30, Date.Parse(sToDate))
            Dim sFromDate As String = dtFromDate.ToString("dd-MMM-yyyy 00:01")
            Call ExportOrders(sFromDate, sToDate)
        Catch ex As Exception
            WebMsgBox.Show("Not a valid date!")
        End Try
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
<head runat="server">
    <title></title>
</head>
<body>
    <form id="Form1" runat="Server">
    <main:Header ID="ctlHeader" runat="server" />
    <table width="100%">
        <tr>
            <td style="width: 1%;">
                &nbsp;
            </td>
            <td style="width: 98%;">
                <asp:Label ID="lblLegendTitle" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                    Font-Bold="True" ForeColor="Gray">Order History</asp:Label>
            </td>
            <td style="width: 1%;">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <asp:Label ID="lblLegendTitle1" runat="server" Font-Size="X-Small" Font-Names="Verdana"
                    Font-Bold="True" ForeColor="Gray">30 Days to:</asp:Label>
            &nbsp;
                <asp:DropDownList ID="ddlToDay" runat="server" Font-Names="Verdana" 
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
                &nbsp;<asp:DropDownList ID="ddlToMonth" runat="server" Font-Names="Verdana" 
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
                </asp:DropDownList>
                &nbsp;<asp:DropDownList ID="ddlToYear" runat="server" Font-Names="Verdana" 
                    Font-Size="XX-Small">
                    <asp:ListItem>2013</asp:ListItem>
                    <asp:ListItem Selected="True">2014</asp:ListItem>
                    <asp:ListItem>2013</asp:ListItem>
                    <asp:ListItem>2015</asp:ListItem>
                    <asp:ListItem>2016</asp:ListItem>
                    <asp:ListItem>2017</asp:ListItem>
                    <asp:ListItem>2018</asp:ListItem>
                    <asp:ListItem>2019</asp:ListItem>
                    <asp:ListItem>2020</asp:ListItem>
                </asp:DropDownList>
            &nbsp;&nbsp;&nbsp;
                <asp:Button ID="btnShow" runat="server" onclick="btnGo_Click" Text="show" 
                    Width="130px" />
            &nbsp;<asp:Button ID="btnExport" runat="server" onclick="btnExport_Click" 
                    Text="export csv" Width="130px" />
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <asp:GridView ID="gvOrders" runat="server" CellPadding="2" Font-Names="Verdana" 
                    Font-Size="XX-Small" Width="100%" OnRowDataBound="gvOrders_RowDataBound">
                </asp:GridView>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
