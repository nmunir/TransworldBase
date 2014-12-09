<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String = "SELECT ProductCode 'Product Code', ProductDescription 'Description', ProductDepartmentId 'Department', CAST(UnitValue AS Numeric(10,2)) 'Cost Price', ProductCategory 'Category', SubCategory 'Sub Category', 'Qty Available' = CASE ISNUMERIC((select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (select sum(LogisticProductQuantity) from LogisticProductLocation AS lpl where lpl.LogisticProductKey = lp.LogisticProductKey) END FROM LogisticProduct lp WHERE DeletedFlag = 'N' AND ArchiveFlag = 'N' AND CustomerKey = 654 ORDER BY ProductCode"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvData.DataSource = dt
        gvData.DataBind()
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
    <title>AAT Product Report</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="gvData" runat="server" CellPadding="2" Width="100%" 
            Font-Names="Arial" Font-Size="Small">
            <AlternatingRowStyle BackColor="#FFFFCC" />
            <RowStyle BackColor="#CCFFFF" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
