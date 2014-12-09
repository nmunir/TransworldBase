<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Globalization" %>
<%@ import Namespace="System.Threading" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")
    Private gsSiteType As String = ConfigLib.GetConfigItem_SiteType
    Private gbSiteTypeDefined As Boolean = gsSiteType.Length > 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Server.ScriptTimeout = 2700
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB", False)

        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call HideAllPanels()
            Call InitWarehouseDropdown("")
            Call InitWarehouseRackDropdown("")
            Call InitWarehouseSectionDropdown("")
            Call InitWarehouseBayDropdown("")

        End If
        Call SetTitle()
    End Sub
    
    Protected Sub HideAllPanels()
        pnlProductList.Visible = False
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Stock by Warehouse Location"
    End Sub
    
    Protected Sub InitWarehouseDropdown(ByVal sWarehouseKey As String)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        If sWarehouseKey = String.Empty Then
            sSQL = "SELECT WarehouseId, WarehouseKey FROM Warehouse ORDER BY WarehouseId"
        Else
            sSQL = "SELECT WarehouseId, WarehouseKey FROM Warehouse WHERE WarehouseKey = " & sWarehouseKey & " ORDER BY WarehouseId"
        End If
        Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
        oConn.Open()
        oDataReader = oCmd.ExecuteReader()
        ddlWarehouse.Items.Add(New ListItem("(all warehouses)", ""))
        While oDataReader.Read
            ddlWarehouse.Items.Add(New ListItem(oDataReader("WarehouseId"), oDataReader("WarehouseKey")))
        End While
        oConn.Close()
    End Sub
    
    Protected Sub InitWarehouseRackDropdown(ByVal sWarehouseKey As String)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        ddlWarehouseRack.Items.Clear()
        ddlWarehouseRack.Items.Add(New ListItem("(all racks)", ""))
        If sWarehouseKey <> String.Empty Then
            sSQL = "SELECT WarehouseRackId, WarehouseRackKey FROM WarehouseRack WHERE WarehouseKey = " & sWarehouseKey & " ORDER BY WarehouseRackId"
            Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                ddlWarehouseRack.Items.Add(New ListItem(oDataReader("WarehouseRackId"), oDataReader("WarehouseRackKey")))
            End While
            oConn.Close()
        End If
    End Sub
    
    Protected Sub InitWarehouseSectionDropdown(ByVal sWarehouseRackKey As String)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        ddlWarehouseSection.Items.Clear()
        ddlWarehouseSection.Items.Add(New ListItem("(all sections)", ""))
        If sWarehouseRackKey <> String.Empty Then
            sSQL = "SELECT WarehouseSectionId, WarehouseSectionKey FROM WarehouseSection WHERE WarehouseRackKey = " & sWarehouseRackKey & " ORDER BY WarehouseSectionId"
            Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                ddlWarehouseSection.Items.Add(New ListItem(oDataReader("WarehouseSectionId"), oDataReader("WarehouseSectionKey")))
            End While
            oConn.Close()
        End If
    End Sub
    
    Protected Sub InitWarehouseBayDropdown(ByVal sWarehouseSectionKey As String)
        Dim oDataReader As SqlDataReader = Nothing
        Dim oConn As New SqlConnection(gsConn)
        Dim sSQL As String
        ddlWarehouseBay.Items.Clear()
        ddlWarehouseBay.Items.Add(New ListItem("(all bays)", ""))
        If sWarehouseSectionKey <> String.Empty Then
            sSQL = "SELECT WarehouseBayId, WarehouseBayKey FROM WarehouseBay WHERE WarehouseSectionKey = " & sWarehouseSectionKey & " ORDER BY WarehouseBayId"
            Dim oCmd As SqlCommand = New SqlCommand(sSQL, oConn)
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            While oDataReader.Read
                ddlWarehouseBay.Items.Add(New ListItem(oDataReader("WarehouseBayId"), oDataReader("WarehouseBayKey")))
            End While
            oConn.Close()
        End If
    End Sub

    Protected Sub ddlWarehouse_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call InitWarehouseRackDropdown(ddl.SelectedValue)
        Call InitWarehouseSectionDropdown("0")
        Call InitWarehouseBayDropdown("0")
        pnlProductList.Visible = False
        'pnlDetail.Visible = False
        'Call ShowProductsByLocation()
    End Sub

    Protected Sub ddlWarehouseRack_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call InitWarehouseSectionDropdown(ddl.SelectedValue)
        Call InitWarehouseBayDropdown("0")
        pnlProductList.Visible = False
        'pnlDetail.Visible = False
        'Call ShowProductsByLocation()
    End Sub

    Protected Sub ddlWarehouseSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        Call InitWarehouseBayDropdown(ddl.SelectedValue)
        pnlProductList.Visible = False
        'pnlDetail.Visible = False
        'Call ShowProductsByLocation()
    End Sub

    Protected Sub ddlWarehouseBay_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        pnlProductList.Visible = False
        'pnlDetail.Visible = False
        'Call ShowProductsByLocation()
    End Sub

    Protected Sub ShowProductsByLocation()
        
    End Sub
    
    Protected Function GenerateQuery() As String
        Dim sbSQL As New StringBuilder
        Dim bListingEverything As Boolean = True
        sbSQL.Append(" SELECT ")
        sbSQL.Append(" CustomerAccountCode 'Customer', lp.ProductCode 'Product Code', LEFT(lp.ProductDescription,20) 'Description', lpl.LogisticProductQuantity 'Qty', wb.WarehouseBayId 'Bay', ws.WarehouseSectionId 'Section', wr.WarehouseRackId 'Rack', w.WarehouseId 'W''house' ")
        sbSQL.Append(" FROM LogisticProduct lp ")
        sbSQL.Append(" INNER JOIN LogisticProductLocation lpl ")
        sbSQL.Append(" ON lp.LogisticProductKey = lpl.LogisticProductKey ")
        sbSQL.Append(" INNER JOIN WarehouseBay AS wb ")
        sbSQL.Append(" ON lpl.WarehouseBayKey = wb.WarehouseBayKey ")
        sbSQL.Append(" INNER JOIN WarehouseSection AS ws ")
        sbSQL.Append(" ON wb.WarehouseSectionKey = ws.WarehouseSectionKey ")
        sbSQL.Append(" INNER JOIN WarehouseRack AS wr ")
        sbSQL.Append(" ON ws.WarehouseRackKey = wr.WarehouseRackKey ")
        sbSQL.Append(" INNER JOIN Warehouse AS w ")
        sbSQL.Append(" ON wr.WarehouseKey = w.WarehouseKey ")
        sbSQL.Append(" INNER JOIN Customer AS c ")
        sbSQL.Append(" ON lp.CustomerKey = c.CustomerKey ")
        If ddlWarehouse.SelectedIndex > 0 Then
            sbSQL.Append(" WHERE w.WarehouseId = '")
            sbSQL.Append(ddlWarehouse.SelectedItem.Text.Replace("'", "''"))
            sbSQL.Append("' ")
            bListingEverything = False
        End If
        If ddlWarehouseRack.SelectedIndex > 0 Then
            sbSQL.Append(" AND wr.WarehouseRackId = '")
            sbSQL.Append(ddlWarehouseRack.SelectedItem.Text.Replace("'", "''"))
            sbSQL.Append("' ")
        End If
        If ddlWarehouseSection.SelectedIndex > 0 Then
            sbSQL.Append(" AND ws.WarehouseSectionId = '")
            sbSQL.Append(ddlWarehouseSection.SelectedItem.Text.Replace("'", "''"))
            sbSQL.Append("' ")
        End If
        If ddlWarehouseBay.SelectedIndex > 0 Then
            sbSQL.Append(" AND wb.WarehouseBayId = '")
            sbSQL.Append(ddlWarehouseBay.SelectedItem.Text.Replace("'", "''"))
            sbSQL.Append("' ")
        End If
        If Not cbIncludeQty0.Checked Then
            If bListingEverything Then
                sbSQL.Append(" WHERE ")
            Else
                sbSQL.Append(" AND ")
            End If
            sbSQL.Append(" lpl.LogisticProductQuantity > 0 ")
        End If

        sbSQL.Append(" ORDER BY ")
        If rbOrderByCustomer.Checked Then
            'sbSQL.Append(" c.CustomerAccountCode, wb.WarehouseBayId, ws.WarehouseSectionId, wr.WarehouseRackId, w.WarehouseId ")
            sbSQL.Append(" c.CustomerAccountCode, w.WarehouseId, wr.WarehouseRackId, ws.WarehouseSectionId, wb.WarehouseBayId ")
        Else
            'sbSQL.Append(" wb.WarehouseBayId, ws.WarehouseSectionId, wr.WarehouseRackId, w.WarehouseId ")
            sbSQL.Append(" w.WarehouseId, wr.WarehouseRackId, ws.WarehouseSectionId, wb.WarehouseBayId ")
        End If
        GenerateQuery = sbSQL.ToString
    End Function
    
    Protected Sub btnListStock_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(GenerateQuery)
        gvProducts.DataSource = oDataTable
        gvProducts.DataBind()
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

    Protected Sub btnExportToExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ExportToExcel()
    End Sub
    
    Protected Sub ExportToExcel()
        Dim sCSVString As String = ConvertDataTableToCSVString(ExecuteQueryToDataTable(GenerateQuery))
        Call ExportCSVData(sCSVString)
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
        Dim s2 As String
        For Each oDataRow In oDataTable.Rows
            For Each s As Object In oDataRow.ItemArray
                Try
                    s2 = s
                Catch
                    s2 = String.Empty
                End Try
                s2 = s2.Replace(Environment.NewLine, " ").Replace("""", "")
                sbResult.Append(s2.Replace(",", " "))
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
   
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Dim sFilename As String = "WH=" & ddlWarehouse.SelectedItem.Text.Replace(" ", "") & "-RACK=" & ddlWarehouseRack.SelectedItem.Text.Replace(" ", "") & "-SECT=" & ddlWarehouseSection.SelectedItem.Text.Replace(" ", "") & "-BAY=" & ddlWarehouseBay.SelectedItem.Text.Replace(" ", "")
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & sFilename & ".csv")
        Response.ContentType = "text/csv"
        'Response.ContentType = "application/vnd.ms-excel"
   
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        Response.End()
        ' Response.Flush()
    End Sub

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        .style1
        {
            width: 15%;
        }
        .style2
        {
            width: 85%;
        }
    </style>
</head>
<body>
    <form id="Form1" runat="Server">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table style="width: 100%" cellpadding="0" cellspacing="0">
            <tr>
                <td style="width: 50%; white-space: nowrap">
                    <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                        Text="Stock by Warehouse Location" Font-Bold="True"/>
                </td>
                <td style="width: 50%; white-space: nowrap" align="right">
                </td>
            </tr>
            <tr>
                <td style="width: 50%; white-space: nowrap">
                    &nbsp;</td>
                <td style="width: 50%; white-space: nowrap" align="right">
                    &nbsp;</td>
            </tr>
        </table>
        <table style="width: 100%" cellpadding="0" cellspacing="0">
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Warehouse:"/>
                </td>
                <td style="width: 95%; white-space: nowrap">
                    <asp:DropDownList ID="ddlWarehouse" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlWarehouse_SelectedIndexChanged" AutoPostBack="True" Height="16px" Width="200px"/>
                </td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Rack:"/>
                </td>
                <td style="width: 95%; white-space: nowrap">
                    <asp:DropDownList ID="ddlWarehouseRack" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlWarehouseRack_SelectedIndexChanged" AutoPostBack="True" Width="200px"/>
                </td>
            </tr>
            <tr>
                <td style="white-space: nowrap" align="right" class="style1">
                    <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Section:"/>
                </td>
                <td style="white-space: nowrap" class="style2">
                    <asp:DropDownList ID="ddlWarehouseSection" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlWarehouseSection_SelectedIndexChanged" AutoPostBack="True" Width="200px"/>
                </td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Bay:"/>
                </td>
                <td style="width: 95%; white-space: nowrap">
                    <asp:DropDownList ID="ddlWarehouseBay" runat="server" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlWarehouseBay_SelectedIndexChanged" AutoPostBack="True" Width="200px"/>
                </td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                        Text="Order by:"/>
                </td>
                <td style="width: 95%; white-space: nowrap">
                    <asp:RadioButton ID="rbOrderByCustomer" runat="server" Checked="True" 
                        Font-Names="Verdana" Font-Size="XX-Small" GroupName="OrderBy" 
                        Text="Customer" />
                    <asp:RadioButton ID="rbOrderByLocation" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" GroupName="OrderBy" Text="Location" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="cbIncludeQty0" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Text="include locations where stock level is zero" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    &nbsp;</td>
                <td style="width: 95%; white-space: nowrap">
                    &nbsp;</td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    &nbsp;</td>
                <td style="width: 95%; white-space: nowrap">
                    <asp:Button ID="btnListStock" runat="server" onclick="btnListStock_Click" 
                        Text="display stock" Width="171px" />
                &nbsp;<asp:Button ID="btnExportToExcel" runat="server" 
                        onclick="btnExportToExcel_Click" Text="export list to excel" 
                        Width="280px" />
                </td>
            </tr>
            <tr>
                <td style="width: 5%; white-space: nowrap" align="right">
                    &nbsp;</td>
                <td style="width: 95%; white-space: nowrap">
                    &nbsp;</td>
            </tr>
            <tr>
                <td style="white-space: nowrap" align="right" colspan="2">
                    <asp:GridView ID="gvProducts" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Width="100%">
                    </asp:GridView>
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlProductList" runat="server" Width="100%">
        </asp:Panel>
    </form>
</body>
</html>

