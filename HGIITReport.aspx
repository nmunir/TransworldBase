<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call GenerateData()
    End Sub
    
    Protected Sub GenerateData()
        Dim sSQL As String = "SELECT hgi.CreatedOn, hgi.AttnOf, hgi.Name, hgi.Addr1, hgi.Addr2, hgi.Town, hgi.Region, hgi.Postcode, hgi.CountryName, hgi.Email, hgi.ConsignmentKey, hgi.OptInOut, hgi.CallerType, hgi.Prompt, hgi.SearchEngine, hgi.Other, hgi.AdvertResponse FROM ClientData_HGI_Order hgi LEFT OUTER JOIN Consignment c ON hgi.ConsignmentKey = c.[key] WHERE ISNULL(CustomerKey, 727) = 727 AND hgi.CreatedOn >= GETDATE()-60 ORDER BY hgi.CreatedOn"
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

        For Each oDataColumn In oDataTable.Columns         ' column headings in line 1
            sbResult.Append(oDataColumn.ColumnName)
            sbResult.Append(",")
        Next
        sbResult.Append(Environment.NewLine)
    
        For Each oDataRow In oDataTable.Rows
            For i As Int32 = 0 To oDataRow.ItemArray.Length - 1
                Dim s2 As String = String.Empty
                If Not IsDBNull(oDataRow.Item(i)) Then
                    s2 = oDataRow.Item(i)
                End If
                s2 = s2.Replace(Environment.NewLine, " ")
                s2 = s2.Replace("""", """""")
                sbResult.Append("""" & s2 & """,")
            Next
            'For Each s As String In oDataRow.ItemArray
            '    Dim s2 As String = String.Empty
            '    If Not IsDBNull(s) Then
            '        s2 = s
            '    End If
            '    s2 = s2.Replace(Environment.NewLine, " ")
            '    s2 = s2.Replace("""", """""")
            '    sbResult.Append("""" & s2 & """,")
            'Next
            sbResult.Length = sbResult.Length - 1
            sbResult.Append(Environment.NewLine)
        Next oDataRow

        If Not sbResult Is Nothing Then
            Return sbResult.ToString()
        Else
            Return String.Empty
        End If
    End Function
    
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "HGIIT_Report_" & Format(Date.Now, "yyyyMMdd") & ".csv")
        Response.ContentType = "text/csv"
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
    <title>Henderson Global Investors - IT Orders - Last 60 Days</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="btnExport" runat="server" OnClick="btnExport_Click" Text="export last 60 days data to excel" />
        <asp:Button ID="btnCloseWindow" runat="server" OnClientClick="javascript: parent.history.back(); return false;" Text="go back" /></div>
    </form>
</body>
</html>