<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const RECORD_TYPE_STORED_PALLET_COUNT As Integer = 1
    Const RECORD_TYPE_MANAGEMENT_FEE As Integer = 2
    Const RECORD_TYPE_PICK_CHARGES As Integer = 3
    Const RECORD_TYPE_SHIPPING_CHARGE As Integer = 4

    Const COLUMN_STORAGE_CHARGE As Integer = 1
    Const COLUMN_MANAGEMENT_FEE As Integer = 2
    Const COLUMN_FULFILMENT As Integer = 3
    Const COLUMN_TOTALS As Integer = 4

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Dim gsMonthName() As String = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
    Dim gnMostRecentYear As Integer
    Dim gnMostRecentMonth As Integer

    Dim gdtDataTable As DataTable
    Dim gdvDataView As DataView
    
    ' NOTE: element 0 is not used in these arrays
    Dim gdblMarketing(4) As Double
    Dim gdblOperations(4) As Double
    Dim gdblTotal(4) As Double
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call InitControls()
            Call GetYears()
            If gnMostRecentYear > 0 Then
                lblMostRecentReport.Text = gsMonthName(gnMostRecentMonth - 1) & " " & gnMostRecentYear.ToString
            Else
                lblMostRecentReport.Text = "... sorry - no report data is available"
                lblMostRecentReport.Font.Bold = True
                lblMostRecentReport.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub
    
    Protected Sub InitTotals()
        For i As Integer = 1 To 4
            gdblMarketing(i) = 0
            gdblOperations(i) = 0
            gdblTotal(i) = 0
        Next
    End Sub
    
    Protected Sub GetYears()
        ddlYear.Items.Add(New ListItem("- please select -", 0))
        For nYear As Integer = 2008 To 2020
            If IsDataForAtLeastOneMonthInYear(nYear) Then
                ddlYear.Items.Add(New ListItem(nYear.ToString, nYear.ToString))
                gnMostRecentYear = nYear
            End If
        Next
    End Sub
    
    Protected Function IsDataForAtLeastOneMonthInYear(ByVal nYear As Integer) As Boolean
        IsDataForAtLeastOneMonthInYear = False
        For nMonth As Integer = 12 To 1 Step -1
            If IsDataForMonth(nYear, nMonth) Then
                IsDataForAtLeastOneMonthInYear = True
                gnMostRecentMonth = nMonth
                Exit For
            End If
        Next
    End Function
    
    Protected Function IsDataForMonth(ByVal nYear As Integer, ByVal nMonth As Integer) As Boolean
        Dim sbSQL As New StringBuilder
        Dim oConn As New SqlConnection(gsConn)
        Dim oDataReader As SqlDataReader = Nothing
        sbSQL.Append("IF EXISTS (SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & nYear.ToString & " AND Month = " & nMonth.ToString & " AND RecordType = 1) ")
        sbSQL.Append("AND EXISTS (SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & nYear.ToString & " AND Month = " & nMonth.ToString & " AND RecordType = 2) ")
        sbSQL.Append("AND EXISTS (SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & nYear.ToString & " AND Month = " & nMonth.ToString & " AND RecordType = 3) ")
        sbSQL.Append("AND EXISTS (SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & nYear.ToString & " AND Month = " & nMonth.ToString & " AND RecordType = 4) ")
        sbSQL.Append("SELECT 1 ELSE SELECT 0")
        Dim oCmd As SqlCommand = New SqlCommand(sbSQL.ToString, oConn)
        IsDataForMonth = False
        Try
            oConn.Open()
            oDataReader = oCmd.ExecuteReader()
            oDataReader.Read()
            If oDataReader(0) = 1 Then
                IsDataForMonth = True
            End If
        Catch ex As Exception
            WebMsgBox.Show("IsDataForMonth: " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
    
    'Protected Function GetRecordsForMonth(ByVal nYear As Integer, ByVal nMonth As Integer) As DataTable
    '    Dim oConn As New SqlConnection(gsConn)
    '    Dim oAdapter As New SqlDataAdapter("SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & nYear.ToString & " AND Month = " & nMonth.ToString, oConn)
    '    gdtDataTable = New DataTable
    '    Try
    '        oAdapter.Fill(gdtDataTable)
    '    Catch ex As Exception
    '        WebMsgBox.Show("GetRecordsForMonth: " & ex.Message)
    '    Finally
    '        oConn.Close()
    '    End Try
    '    GetRecordsForMonth = gdtDataTable
    'End Function

    Protected Sub GetMonths()
        ddlMonth.Items.Clear()
        ddlMonth.Items.Add(New ListItem("- please select -", 0))
        For nMonth As Integer = 1 To 12
            If IsDataForMonth(ddlYear.SelectedValue, nMonth) Then
                ddlMonth.Items.Add(New ListItem(gsMonthName(nMonth - 1), nMonth))
                ddlMonth.Visible = True
                lblLegendMonth.Visible = True
                Exit For
            End If
        Next
    End Sub
    
    Protected Sub InitControls()
        lblLegendMonth.Visible = False
        ddlMonth.Visible = False
        btnDownload.Visible = False
    End Sub
    
    Protected Sub ddlYear_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlYear.Items(0).Value = 0 Then
            ddlYear.Items.RemoveAt(0)
            lblLegendMonth.Visible = True
            ddlMonth.Visible = True
        End If
        Call GetMonths()
        btnDownload.Visible = False
    End Sub
    
    Protected Sub ddlMonth_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlMonth.Items(0).Value = 0 Then
            ddlMonth.Items.RemoveAt(0)
            btnDownload.Visible = True
        End If
    End Sub
    
    Private Sub ExportCSVData(ByVal sCSVString As String)
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" & "FEXCOCostByCategory " & gsMonthName(ddlMonth.SelectedValue - 1) & " " & ddlYear.SelectedValue.ToString & ".csv")
        'Response.ContentType = "application/vnd.ms-excel"
        Response.ContentType = "text/csv"
    
        Dim eEncoding As Encoding = Encoding.GetEncoding("Windows-1252")
        Dim eUnicode As Encoding = Encoding.Unicode
        Dim byUnicodeBytes As Byte() = eUnicode.GetBytes(sCSVString)
        Dim byEncodedBytes As Byte() = Encoding.Convert(eUnicode, eEncoding, byUnicodeBytes)
        Response.BinaryWrite(byEncodedBytes)
        'Response.Flush()
        Response.End()
    End Sub

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitTotals()
        gdtDataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_FEXCO_AllocatedCharges WHERE Year = " & ddlYear.SelectedValue & " AND Month = " & ddlMonth.SelectedValue)
        'Call GetRecordsForMonth(ddlYear.SelectedValue, ddlMonth.SelectedValue)
        gdvDataView = New DataView(gdtDataTable)
        Call ExtractTotalsManagementFee()
        Call ExtractTotalsStoredPallets()
        Call CalculatePalletStorageCharge()
        Call ExtractTotalsFulfilment()
        Call CalculateGrandTotals()
        Call ExportCSVData(BuildCSVString())
    End Sub
    
    Protected Sub AppendData(ByVal dr As DataRow, ByVal RecordType As Integer)
        If Not IsDBNull(dr("Marketing")) Then
            gdblMarketing(RecordType) = gdblMarketing(RecordType) + dr("Marketing")
        End If
        If Not IsDBNull(dr("Operations")) Then
            gdblOperations(RecordType) = gdblOperations(RecordType) + dr("Operations")
        End If
    End Sub
        
    Protected Sub ExtractTotalsManagementFee()
        gdvDataView.RowFilter = "RecordType=" & RECORD_TYPE_MANAGEMENT_FEE
        If gdvDataView.Count = 1 Then
            Call AppendData(gdvDataView.Item(0).Row, COLUMN_MANAGEMENT_FEE)
        Else
            WebMsgBox.Show("ExtractTotalsManagementFee: WARNING - unexpected number of records encountered. Please inform your account handler.")
        End If
    End Sub
    
    Protected Sub ExtractTotalsStoredPallets()
        gdvDataView.RowFilter = "RecordType=" & RECORD_TYPE_STORED_PALLET_COUNT
        If gdvDataView.Count = 1 Then
            Call AppendData(gdvDataView.Item(0).Row, COLUMN_STORAGE_CHARGE)
        Else
            WebMsgBox.Show("ExtractTotalsStoredPallets: WARNING - unexpected number of records encountered. Please inform your account handler.")
        End If
    End Sub

    Protected Sub CalculatePalletStorageCharge()
        gdblMarketing(COLUMN_STORAGE_CHARGE) = gdblMarketing(COLUMN_STORAGE_CHARGE) * pdblPalletMonthlyStorageCost
        gdblOperations(COLUMN_STORAGE_CHARGE) = gdblOperations(COLUMN_STORAGE_CHARGE) * pdblPalletMonthlyStorageCost
    End Sub
        
    Protected Sub ExtractTotalsFulfilment()
        gdvDataView.RowFilter = "RecordType=" & RECORD_TYPE_PICK_CHARGES
        If gdvDataView.Count > 0 Then
            For Each drv As DataRowView In gdvDataView
                Call AppendData(drv.Row, COLUMN_FULFILMENT)
            Next
        Else
            WebMsgBox.Show("ExtractTotalsFulfilment: WARNING - no pick fee records found. Please inform your account handler.")
        End If

        gdvDataView.RowFilter = "RecordType=" & RECORD_TYPE_SHIPPING_CHARGE
        If gdvDataView.Count > 0 Then
            For Each drv As DataRowView In gdvDataView
                Call AppendData(drv.Row, COLUMN_FULFILMENT)
            Next
        Else
            WebMsgBox.Show("ExtractTotalsFulfilment: WARNING - no shipping charge records found. Please inform your account handler.")
        End If
    End Sub

    Protected Sub CalculateGrandTotals()
        For i As Integer = COLUMN_STORAGE_CHARGE To COLUMN_FULFILMENT
            gdblMarketing(COLUMN_TOTALS) = gdblMarketing(COLUMN_TOTALS) + gdblMarketing(i)
            gdblOperations(COLUMN_TOTALS) = gdblOperations(COLUMN_TOTALS) + gdblOperations(i)
            
            gdblTotal(i) = gdblTotal(i) + gdblMarketing(i)
            gdblTotal(i) = gdblTotal(i) + gdblOperations(i)
        Next
        gdblTotal(COLUMN_TOTALS) = gdblTotal(COLUMN_STORAGE_CHARGE) + gdblTotal(COLUMN_MANAGEMENT_FEE) + gdblTotal(COLUMN_FULFILMENT)
    End Sub

    Protected Function BuildCSVString() As String
        Dim sbCSV As New StringBuilder
        sbCSV.Append("Cost by Category for " & gsMonthName(ddlMonth.SelectedValue - 1) & " " & ddlYear.SelectedValue.ToString & ",,,," & Environment.NewLine)
        sbCSV.Append(",,,," & Environment.NewLine)
        sbCSV.Append("Product Category, Management, Storage, Fulfilment/Handling, TOTAL" & Environment.NewLine)

        sbCSV.Append("Marketing,")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblMarketing(COLUMN_MANAGEMENT_FEE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblMarketing(COLUMN_STORAGE_CHARGE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblMarketing(COLUMN_FULFILMENT))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblMarketing(COLUMN_TOTALS))))
        sbCSV.Append(Environment.NewLine)
        
        sbCSV.Append("Operations,")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblOperations(COLUMN_MANAGEMENT_FEE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblOperations(COLUMN_STORAGE_CHARGE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblOperations(COLUMN_FULFILMENT))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblOperations(COLUMN_TOTALS))))
        sbCSV.Append(Environment.NewLine)
        
        sbCSV.Append("TOTAL,")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblTotal(COLUMN_MANAGEMENT_FEE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblTotal(COLUMN_STORAGE_CHARGE))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblTotal(COLUMN_FULFILMENT))))
        sbCSV.Append(",")
        sbCSV.Append(sQuotedString(String.Format("{0:c}", gdblTotal(COLUMN_TOTALS))))
        sbCSV.Append(Environment.NewLine)
        
        BuildCSVString = sbCSV.ToString
    End Function

    Protected Function sQuotedString(ByVal sTemp) As String
        sQuotedString = """" & sTemp & """"
    End Function

    Protected Function ExecuteNonQuery(ByVal sQuery As String) As Boolean
        ExecuteNonQuery = False
        Dim oConn As New SqlConnection(gsConn)
        Dim oCmd As SqlCommand
        Try
            oConn.Open()
            oCmd = New SqlCommand(sQuery, oConn)
            oCmd.ExecuteNonQuery()
            ExecuteNonQuery = True
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteNonQuery executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
    End Function
   
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

    Property pdblPalletMonthlyStorageCost() As Double
        Get
            Dim o As Object = ViewState("FR_PalletStorageCost")
            If o Is Nothing Then
                Dim oDataTable As DataTable = ExecuteQueryToDataTable("SELECT PalletWeeklyFee FROM ClientData_FEXCO_Configuration")
                Dim dbl As Double = CDbl(oDataTable.Rows(0)(0)) * 4  ' do cost multiplication here since cost is weekly and report is monthly
                ViewState("FR_PalletStorageCost") = dbl
                Return dbl
            End If
            Return CDbl(o)
        End Get
        Set(ByVal Value As Double)
            ViewState("FR_PalletStorageCost") = Value
        End Set
    End Property

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Western Union Retail Services GB Cost By Category Report</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        &nbsp;<asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana"
            Font-Size="X-Small" Text="Western Union Retail Services GB Cost By Category Report"></asp:Label><br />
        <br />
        &nbsp;<asp:Label ID="lblMostRecentReportPreamble" runat="server" Font-Names="Verdana"
            Font-Size="XX-Small" Text="The most recent report for which full data appears to be available is "></asp:Label>
        <asp:Label ID="lblMostRecentReport" runat="server" Font-Names="Verdana" Font-Size="XX-Small"></asp:Label>
        <br />
        <br />
        &nbsp;<asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Select year and month"></asp:Label><br />
        <br />
        &nbsp;<asp:Label ID="lblLegendYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
            Text="Year:"></asp:Label>
        <asp:DropDownList ID="ddlYear" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlYear_SelectedIndexChanged">
        </asp:DropDownList>
        &nbsp;
        <asp:Label ID="lblLegendMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Month:"></asp:Label>&nbsp;<asp:DropDownList
            ID="ddlMonth" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlMonth_SelectedIndexChanged" Visible="False">
        </asp:DropDownList>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;<asp:Button ID="btnDownload" runat="server" Text="download to excel" Visible="False" OnClick="btnDownload_Click" />
        &nbsp; &nbsp; &nbsp;
        <asp:Button ID="btnCloseWindow" runat="server" OnClientClick="window.close()" Text="close window" />
    </div>
    </form>
</body>
</html>
