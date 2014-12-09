<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register assembly="Telerik.Web.UI" namespace="Telerik.Web.UI" tagprefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' implement month dropdown
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        Server.ScriptTimeout = 1000
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("./session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call GetCustomerAccountCodes()
        End If
        'Call GetCustomerAccountCodes()
    End Sub
    
    Protected Sub GetCustomerAccountCodes()
        Dim sSQL As String = "SELECT DISTINCT c.CustomerAccountCode, lp.CustomerKey FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE c.DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim lic As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        rcbCustomers.Items.Clear()
        rcbCustomers.Items.Add(New RadComboBoxItem("- please select - ", 0))
        For Each li As ListItem In lic
            rcbCustomers.Items.Add(New RadComboBoxItem(li.Text, li.Value))
        Next
    End Sub

    Protected Sub btnGenerateReport_Click(sender As Object, e As System.EventArgs)
        If rdpStartDate.IsEmpty Or rdpEndDate.IsEmpty Then
            WebMsgBox.Show("Please ensure you ahve entered a valid enter start and end date.")
            Exit Sub
        End If
        If Not rcbCustomers.SelectedIndex > 0 Then
            WebMsgBox.Show("Please select the customer for whom you want to generate the product report.")
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
            oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@FromDate", SqlDbType.SmallDateTime))
            dtDate = rdpStartDate.SelectedDate
            sStartDate = Format(dtDate, "d-MMM-yyyy")
            oAdapter.SelectCommand.Parameters("@FromDate").Value = sStartDate
            
            dtDate = rdpEndDate.SelectedDate
            sEndDate = Format(dtDate, "d-MMM-yyyy")
            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ToDate", SqlDbType.SmallDateTime))
            oAdapter.SelectCommand.Parameters("@ToDate").Value = sEndDate

            oAdapter.SelectCommand.Parameters.Add(New SqlParameter("@CustomerKey", SqlDbType.Int))
            oAdapter.SelectCommand.Parameters("@CustomerKey").Value = rcbCustomers.SelectedValue
            
            oAdapter.SelectCommand.CommandTimeout = 0
            
            oAdapter.Fill(dt)
            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append(rcbCustomers.SelectedItem.Text)
            sResponseValue.Append("_")
            sResponseValue.Append(sStartDate.Replace("-", ""))
            sResponseValue.Append("-")
            sResponseValue.Append(sEndDate.Replace("-", ""))
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)
            Response.Write("Product Code, Value/Date, Description, Supplier, " & sStartDate & " Qty, " & sEndDate & " Qty, Despatched Qty, Return Qty, Goods In Qty, Adjustment Qty, Cost Price, Sales Price" & vbNewLine)
            Dim sItem As String
            For Each dr As DataRow In dt.Rows
                For i = 2 To dt.Columns.Count - 1
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
            WebMsgBox.Show("Failed executing stored procedure spASPNET_ProductQuantitiesReport2 " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Protected Function ExecuteQueryToListItemCollection(ByVal sQuery As String, ByVal sTextFieldName As String, ByVal sValueFieldName As String) As ListItemCollection
        Dim oListItemCollection As New ListItemCollection
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sTextField As String
        Dim sValueField As String
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    If Not IsDBNull(oDataReader(sTextFieldName)) Then
                        sTextField = oDataReader(sTextFieldName)
                    Else
                        sTextField = String.Empty
                    End If
                    If Not IsDBNull(oDataReader(sValueFieldName)) Then
                        sValueField = oDataReader(sValueFieldName)
                    Else
                        sValueField = String.Empty
                    End If
                    oListItemCollection.Add(New ListItem(sTextField, sValueField))
                End While
            End If
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

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
    
        &nbsp;<asp:Label ID="Label1" runat="server" Text="Product Report Generator"/>
    
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
                    <asp:Label ID="Label2" runat="server" Text="Month:"/>
                </td>
                <td>
                    <telerik:RadComboBox ID="rcbMonth" Runat="server" Enabled="False">
                    </telerik:RadComboBox>
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
                    <telerik:RadComboBox ID="rcbCustomers" Runat="server">
                    </telerik:RadComboBox>
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
