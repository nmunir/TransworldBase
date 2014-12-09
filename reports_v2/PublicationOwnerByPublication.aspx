<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString()
    Private gsSQL As String = "SELECT ProductCode 'Product Code', ISNULL(ProductDate,'') 'Value / Date', ProductDescription 'Description', FirstName + ' ' + LastName + ' (' + UserId + ')' 'Owner', CONVERT(VARCHAR(11), lp.CreatedOn, 106) 'Created On' FROM LogisticProduct lp INNER JOIN UserProfile up ON lp.StockOwnedByKey =  up.[key] WHERE lp.CustomerKey = 663 AND ISNULL(lp.StockOwnedByKey,0) > 0 AND lp.DeletedFlag = 'N' ORDER BY lp.ProductCode"

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("../session_expired.aspx")
        End If
    End Sub
    
    Protected Sub btnShowPublications_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        gvPublications.DataSource = oDataTable
        gvPublications.DataBind()
    End Sub
    
    Protected Sub btnExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExportToExcel()
    End Sub
    
    Protected Sub ExportToExcel()
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(gsSQL)
        If oDataTable.Rows.Count > 0 Then
            Response.Clear()
            Response.ContentType = "text/csv"

            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append("PublicationsByOwner.csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)

            For Each c As DataColumn In oDataTable.Columns
                Response.Write(c.ColumnName)
                Response.Write(",")
            Next
            Response.Write(vbCrLf)
            For Each r As DataRow In oDataTable.Rows
                Dim dr As DataRow = r
                For Each sItem As String In dr.ItemArray
                    sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        Else
            WebMsgBox.Show("No data found.")
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
<head runat="server">
    <title>Publication Owner By Publication</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;
        <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" Text="Show Publication Owner By Publication"/>
        <br />
        <br />
        &nbsp;<asp:Button ID="btnShowPublications" runat="server" Text="show publications" onclick="btnShowPublications_Click" />
        &nbsp;<asp:Button ID="btnExportToExcel" runat="server" Text="export publications list to excel" onclick="btnExportToExcel_Click" />
        <br />
        <br />
        <asp:GridView ID="gvPublications" runat="server" CellPadding="2" 
            Font-Names="Verdana" Font-Size="XX-Small" Width="100%">
            <AlternatingRowStyle BackColor="#EEEEEE" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
