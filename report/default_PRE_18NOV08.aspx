<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>

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
        Dim sSQL As String = "SELECT * FROM ClientData_DrinkAware WHERE EntryTimestamp >= '" & dtStart.ToString("dd-MMM-yyyy") & "' ORDER BY EntryTimestamp"
        Dim oConn As New SqlConnection(gsConn)
        Try
            Dim oAdapter As New SqlDataAdapter(sSQL, oConn)
            Dim oDataTable As New DataTable
            oAdapter.Fill(oDataTable)
            Dim sCSVString As String = ConvertDataTableToCSVString2(oDataTable)
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

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            sbResult.Length = sbResult.Length - 1
        End If
        sbResult.Append(Environment.NewLine)
    
        For Each oDataRow In oDataTable.Rows
            For Each s As String In oDataRow.ItemArray
                Dim s2 As String = s.Replace(Environment.NewLine, " ")
                sbResult.Append(s2.Replace(",", " "))
                sbResult.Append(",")
            Next
            'For Each oDataColumn In oDataTable.Columns
            '    sbResult.Append(oDataRow(Replace(oDataColumn.ColumnName, ",", " ")))  ' replace any commas with spaces
            '    sbResult.Append(",")
            'Next oDataColumn
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
    
    Public Function ConvertDataTableToCSVString2(ByVal oDataTable As DataTable) As String
        Dim sbResult As New StringBuilder
        Dim oDataColumn As DataColumn
        Dim oDataRow As DataRow

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        If sbResult.Length > 1 Then
            'sbResult.Length = sbResult.Length - 1
            sbResult.Append("Order")
        End If
        sbResult.Append(Environment.NewLine)
    
        For Each oDataRow In oDataTable.Rows
            For Each s As String In oDataRow.ItemArray
                Dim s2 As String = s.Replace(Environment.NewLine, " ")
                sbResult.Append(s2.Replace(",", " "))
                sbResult.Append(",")
            Next
            'For Each oDataColumn In oDataTable.Columns
            '    sbResult.Append(oDataRow(Replace(oDataColumn.ColumnName, ",", " ")))  ' replace any commas with spaces
            '    sbResult.Append(",")
            'Next oDataColumn
            'sbResult.Length = sbResult.Length - 1
            sbResult.Append(GetOrderDetails(oDataRow("ConsignmentKey")))
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
    
    Protected Function GetOrderDetails(ByVal sConsignmentNumber As String) As String
        Dim sbOrderDetails As New StringBuilder
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        Dim sSQL As String = "SELECT DISTINCT DistributionListName FROM AddressDistributionLists WHERE CustomerKey = " & Session("CustomerKey") & " ORDER BY DistributionListName"
        Dim oCmd As SqlCommand = New SqlCommand("spASPNET_LogisticBooking_GetMovementsWithVals", oConn)
        oCmd.CommandType = CommandType.StoredProcedure
        Dim oParam As SqlParameter = oCmd.Parameters.Add("@ConsignmentKey", SqlDbType.Int)
        oParam.Value = CInt(sConsignmentNumber)
        
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                Dim sOrderDetails As String = oDataReader("ProductCode") & " - " & oDataReader("ItemsOut") & ";"
                sbOrderDetails.Append(sOrderDetails)
            End While
        Catch ex As Exception
            WebMsgBox.Show("GetOrderDetails: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        GetOrderDetails = sbOrderDetails.ToString
    End Function
    
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "DrinkAware Demographic Report.csv")
        Response.ContentType = "application/vnd.ms-excel"
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)

        Response.Flush()

        ' Stop execution of the current page
        'Response.End()
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