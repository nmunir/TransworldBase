<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register assembly="Telerik.Web.UI" namespace="Telerik.Web.UI" tagprefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' implement month dropdown

    Const CUSTOMER_HYSTER As String = "77"
    Const CUSTOMER_YALE As String = "680"
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Server.ScriptTimeout = 1000
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("./session_expired.aspx")
        End If
        If Not IsPostBack Then
        End If
    End Sub
    
    Protected Sub btnGenerateReport_Click(sender As Object, e As System.EventArgs)
        If rdpStartDate.IsEmpty Or rdpEndDate.IsEmpty Then
            WebMsgBox.Show("Please ensure you ahve entered a valid enter start and end date.")
            Exit Sub
        End If
        If rdpStartDate.SelectedDate >= rdpEndDate.SelectedDate Then
            WebMsgBox.Show("Start date must be before End Date. Please re-select the date range.")
            Exit Sub
        End If
        Call GenerateReport()
    End Sub
    
    Protected Sub GenerateReport()
        Dim dt As New DataTable
        Dim dtDate As Date
        Dim sStartDate As String, sEndDate As String
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_ProductQuantitiesReport3", oConn)
        Try
            dtDate = rdpStartDate.SelectedDate
            sStartDate = Format(dtDate, "d-MMM-yyyy") & " 00:01"
            
            dtDate = rdpEndDate.SelectedDate
            sEndDate = Format(dtDate, "d-MMM-yyyy") & " 23:59"
            Dim sCustomerClause As String = " AND c.CustomerKey IN ("
            If cbHyster.Checked Then
                sCustomerClause += CUSTOMER_HYSTER
            End If
            If cbYale.Checked Then
                If cbHyster.Checked Then
                    sCustomerClause += ", "
                End If
                sCustomerClause += CUSTOMER_YALE
            End If
            sCustomerClause += ") "
            
            Dim sSQL As String = "SELECT ISNULL(CAST(REPLACE(CONVERT(VARCHAR(11), c.CreatedOn, 106), ' ', '-') AS varchar(20)),'(never)') 'CreatedOn', hycce.ConsignmentKey 'AWB', cust.CustomerAccountCode, ServiceLevel, EstimatedWeight, EstimatedCost, CneeCtcName 'ContactName', CneeName 'Name', CneeAddr1 'Addr1', CneeAddr2 'Addr2', CneeTown 'Town', CneePostcode 'Postcode', ItemsOut 'Quantity', lp.ProductCode 'ProductCode', lp.ProductDescription 'Description' FROM ClientData_HysterYale_ConsignmentCostEstimate hycce INNER JOIN Consignment c ON hycce.ConsignmentKey = c.[key] INNER JOIN LogisticMovement lm ON c.[key] = lm.ConsignmentKey INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey INNER JOIN Customer cust ON c.CustomerKey = cust.CustomerKey WHERE c.CreatedOn >= '" & sStartDate & "' AND c.CreatedOn <= '" & sEndDate & "'" & sCustomerClause & " ORDER BY hycce.[id]"
            dt = ExecuteQueryToDataTable(sSQL)

            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append("HysterYaleOrderReport")
            sResponseValue.Append("_")
            sResponseValue.Append(sStartDate.Replace("-", ""))
            sResponseValue.Append("-")
            sResponseValue.Append(sEndDate.Replace("-", ""))
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)
            Response.Write("Created On, AWB, Customer, Service Level, Est. Weight, Est. Cost, Contact Name, Name, Addr 1, Addr 2, Town, Postcode, Quantity, Product Code, Description" & vbNewLine)
            Dim sItem As String
            For Each dr As DataRow In dt.Rows
                For i = 0 To dt.Columns.Count - 1
                    sItem = dr(i) & String.Empty
                    sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Catch ex As SqlException
            WebMsgBox.Show("Failed to retrieve data for Hyster/Yale Orders report: " & ex.Message)
        Finally
            oConn.Close()
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
    <form id="form1" runat="server">
    <main:Header id="ctlHeader" runat="server"></main:Header>
    <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
    
        &nbsp;<asp:Label ID="Label1" runat="server" Text="Hyster / Yale Orders - Cost Estimates"/>
    
        <table style="width: 100%">
            <tr>
                <td style="width: 5%">
                    &nbsp;</td>
                <td style="width: 20%">
                    &nbsp;</td>
                <td style="width: 20%">
                    &nbsp;</td>
                <td style="width: 20%">
                    &nbsp;</td>
                <td style="width: 20%">
                    &nbsp;</td>
                <td style="width: 15%">
                    &nbsp;</td>
            </tr>
            <tr>
                <td/>
                <td align="right">
                    <asp:Label ID="Label3" runat="server" Text="Start Date:"/>
                </td>
                <td>
                    <telerik:RadDatePicker ID="rdpStartDate" Runat="server">
                    </telerik:RadDatePicker>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td/>
                <td align="right">
                    <asp:Label ID="Label4" runat="server" Text="End Date:"/>
                </td>
                <td>
                    <telerik:RadDatePicker ID="rdpEndDate" Runat="server">
                    </telerik:RadDatePicker>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td/>
                    &nbsp;<td align="right">
                    <asp:Label ID="Label5" runat="server" Text="Customer:"/>
                </td>
                <td>
                    <asp:CheckBox ID="cbHyster" runat="server" Checked="True" Text="Hyster" />
&nbsp;<asp:CheckBox ID="cbYale" runat="server" Checked="True" Text="Yale" />
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td/>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td/>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnGenerateReport" runat="server" 
                        onclick="btnGenerateReport_Click" Text="Generate Report" />
                </td>
                <td>
                </td>
                <td>
                </td>
                <td>
                </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
