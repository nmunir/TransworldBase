<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const PERIOD As Integer = 90
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GenerateData()
    End Sub
    
    Protected Sub GenerateData()
        Dim dtEnd As Date = DateAdd(DateInterval.Day, 1, Now())
        Dim dtStart As Date = DateAdd(DateInterval.Day, -(PERIOD + 1), dtEnd)
        'Dim sSQL As String = "SELECT * FROM ClientData_DrinkAware WHERE EntryTimestamp >= '" & dtStart.ToString("dd-MMM-yyyy") & "' ORDER BY EntryTimestamp"
        Dim sSQL As String = "SELECT id 'Record #', Title, Name, Company, JobTitle 'Job Title', Address, Postcode, Email, Telephone, Comments, TypeOfOrganisation 'Type of Organisation', Optin 'Opt In', ConsignmentKey 'Consignment No.', EntryTimestamp 'Order Placed', PurchaseOrderNo 'Purchase Order No.', Mode 'User Type' FROM ClientData_DrinkAware WHERE EntryTimestamp >= '" & dtStart.ToString("dd-MMM-yyyy") & "' ORDER BY EntryTimestamp"
        Dim oConn As New SqlConnection(gsConn)
        Try
            Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
            Dim oDataTable As New DataTable
            oAdapter.Fill(oDataTable)
            Dim sCSVString As String = ConvertDataTableToCSVString(oDataTable)
            Call ExportCSVData(sCSVString)
        Catch ex As Exception
            WebMsgBox.Show("Error in GenerateData: " & ex.Message)
        Finally
            oConn.Close()
            Response.End()
        End Try
    End Sub

    Public Function ConvertDataTableToCSVString(ByVal oDataTable As DataTable) As String
        Dim sbResult As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRow
        Dim dictProductsUsed As Dictionary(Of String, String)
        Dim dictProductRow As New Dictionary(Of String, Integer)

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            dictProductsUsed = GetProductsUsed()
            For Each kv As KeyValuePair(Of String, String) In dictProductsUsed   ' assume (a) there is at least one product used in the last n days, (b) product codes do not contain quote marks or commas
                sbResult.Append(kv.Value.Replace(",", " "))
                sbResult.Append(",")
                dictProductRow.Add(kv.Key, 0)
            Next
            sbResult.Length = sbResult.Length - 1
        End If
        sbResult.Append(Environment.NewLine)
    
        For Each oDataRow In oDataTable.Rows
            For Each s As Object In oDataRow.ItemArray
                Dim sX As String = s & String.Empty
                Dim s2 As String = sX.Replace(Environment.NewLine, " ")
                sbResult.Append(s2.Replace(",", " "))
                sbResult.Append(",")
            Next
            Call GetOrderDetails(oDataRow("Consignment No."), dictProductRow)
            For Each kv As KeyValuePair(Of String, Integer) In dictProductRow
                sbResult.Append(kv.Value)
                sbResult.Append(",")
            Next
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
    
    Protected Function GetProductsUsed() As Dictionary(Of String, String)
        Dim dictProductsUsed As New Dictionary(Of String, String)
        GetProductsUsed = dictProductsUsed
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT ProductCode, ProductCode + ' ' + SUBSTRING(ProductDescription,0,100) ")
        sbSQL.Append("FROM LogisticProduct ")
        sbSQL.Append("WHERE LogisticProductKey IN ")
        sbSQL.Append("(SELECT DISTINCT lm.LogisticProductKey ")
        sbSQL.Append("FROM	LogisticProduct AS lp ")
        sbSQL.Append("LEFT OUTER JOIN LogisticMovement AS lm ")
        sbSQL.Append("ON lp.LogisticProductKey = lm.LogisticProductKey ")
        sbSQL.Append("WHERE lm.CustomerKey = 546 ")
        sbSQL.Append("AND ItemsOut > 0 ")
        sbSQL.Append("AND LogisticMovementStartDateTime > (GETDATE() - 90)) ")
        sbSQL.Append("ORDER BY ProductCode")
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            If oDataReader.HasRows Then
                While oDataReader.Read
                    dictProductsUsed.Add(oDataReader(0), oDataReader(1))
                End While
            End If
            GetProductsUsed = dictProductsUsed
        Catch ex As Exception
            WebMsgBox.Show("GetProductsUsed: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    Protected Sub GetOrderDetails(ByVal sConsignmentNumber As String, ByRef dictProductsUsed As Dictionary(Of String, Integer))
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_GetMovementsWithVals", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@ConsignmentKey", SqlDbType.Int)
        oParam.Value = CInt(sConsignmentNumber)
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            Dim arrlst As New ArrayList
            For Each s As String In dictProductsUsed.Keys
                arrlst.Add(s)
            Next
            For i As Integer = 0 To arrlst.Count - 1
                dictProductsUsed(arrlst(i)) = 0
            Next
            While oDataReader.Read
                dictProductsUsed(oDataReader("ProductCode")) = oDataReader("ItemsOut")
            End While
        Catch ex As Exception
            WebMsgBox.Show("GetOrderDetails: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Sub
    
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "DrinkAware_Demographic_Report.csv")
        Response.ContentType = "application/vnd.ms-excel"
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()
    End Sub
    

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>DrinkAware Orders Demographic Data - Last 90 Days</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="btnExport" runat="server" OnClick="btnExport_Click" Text="export last 90 days data to excel" />
        <asp:Button ID="btnCloseWindow" runat="server" OnClientClick="window.close()" Text="close window" /></div>
    </form>
</body>
</html>