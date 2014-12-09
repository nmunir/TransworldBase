<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const CUSTOMER_JAMESHARDIE As Int32 = 837
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Call GenerateData()
        End If
    End Sub

    Protected Sub GenerateData()
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter("spASPNET_Report_JamesHardieOrdersCurrentMonthUsersByValue01", oConn)
        oAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        Try
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("spASPNET_Report_JamesHardieOrdersCurrentMonthUsersByValue01: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        gvResults.DataSource = oDataTable
        gvResults.DataBind()
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
            ExecuteQueryToDataTable = Nothing
        Finally
            oConn.Close()
            ExecuteQueryToDataTable = oDataTable
        End Try
    End Function

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>James Hardie Orders - Current Month - Users By Value</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="gvResults" runat="server" CellPadding="2" Font-Names="Verdana"
            Font-Size="Small">
        </asp:GridView>
    </div>
    </form>
</body>
</html>