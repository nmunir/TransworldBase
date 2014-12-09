<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.IO" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>
<%@ import Namespace="System.Collections.Generic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    Dim goControl As Control = Nothing
    Dim garrMonths() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call HideAllPanels()
            pnlCustomer.Visible = True
            Call BindTariffsGrid()
        End If
        Call SetTitle()
        'tbSearch.Attributes.Add("onkeypress", "return clickButton(event,'" + btnSearchGo.ClientID + "')")
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Consumables"
    End Sub
    
    Protected Sub HideAllPanels()
        pnlCustomer.Visible = False
        pnlDetail.Visible = False
    End Sub
    
    Protected Sub ShowCustomerPanel()
        Call HideAllPanels()
        pnlCustomer.Visible = True
    End Sub
    
    Protected Sub ShowDetailPanel()
        Call HideAllPanels()
        pnlDetail.Visible = True
    End Sub
    
    Protected Sub PopulateCustomerDropdown()
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection("SELECT CustomerName + ' (' + CustomerAccountCode + ')' CustName, CustomerAccountCode FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND ISNULL(AccountHandlerKey, 0) > 0 AND NOT CustomerKey IN (SELECT CustomerKey FROM ConsumablesCost) ORDER BY CustomerAccountCode", "CustName", "CustomerAccountCode")
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
    End Sub
    
    Protected Sub BindTariffsGrid()
        Dim sbSQL As New StringBuilder
        sbSQL.Append("SELECT ")
        sbSQL.Append("'AAAAAAAA' AS 'Customer' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A4Jiffy AS DECIMAL(8,2)) 'A4 Jiffy' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A5Jiffy AS DECIMAL(8,2)) 'A5 Jiffy' ")
        sbSQL.Append(",")
        sbSQL.Append("CAST(A3Jiffy AS DECIMAL(8,2)) 'A3 Jiffy' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A4Box AS DECIMAL(8,2)) 'A4 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A5Box AS DECIMAL(8,2)) 'A5 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A3Box AS DECIMAL(8,2)) 'A3 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(SmallTube AS DECIMAL(8,2)) 'Small Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(MedTube AS DECIMAL(8,2)) 'Medium Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(LargeTube AS DECIMAL(8,2)) 'Large Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(FTC AS DECIMAL(8,2)) 'FTC' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(Infill AS DECIMAL(8,2)) 'Infill' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(Other AS DECIMAL(8,2)) 'Other' ")
        sbSQL.Append("FROM ConsumablesCost cc ")
        sbSQL.Append("WHERE CustomerKey = 0 ")

        sbSQL.Append("UNION ")
        
        sbSQL.Append("SELECT ")
        sbSQL.Append("CustomerName + ' ~ ' + CustomerAccountCode AS 'Customer' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A4Jiffy AS DECIMAL(8,2)) 'A4 Jiffy' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A5Jiffy AS DECIMAL(8,2)) 'A5 Jiffy' ")
        sbSQL.Append(",")
        sbSQL.Append("CAST(A3Jiffy AS DECIMAL(8,2)) 'A3 Jiffy' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A4Box AS DECIMAL(8,2)) 'A4 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A5Box AS DECIMAL(8,2)) 'A5 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(A3Box AS DECIMAL(8,2)) 'A3 Box' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(SmallTube AS DECIMAL(8,2)) 'Small Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(MedTube AS DECIMAL(8,2)) 'Medium Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(LargeTube AS DECIMAL(8,2)) 'Large Tube' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(FTC AS DECIMAL(8,2)) 'FTC' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(Infill AS DECIMAL(8,2)) 'Infill' ")
        sbSQL.Append(", ")
        sbSQL.Append("CAST(Other AS DECIMAL(8,2)) 'Other' ")
        sbSQL.Append("FROM ConsumablesCost cc ")
        sbSQL.Append("INNER JOIN Customer c ")
        sbSQL.Append("ON cc.CustomerKey = c.CustomerKey ")
        sbSQL.Append("ORDER BY Customer")
        'sbSQL.Append("")
        Dim sSQL As String = sbSQL.ToString
        Dim oDataTable As DataTable = ExecuteQueryToDataTable(sSQL)
        gvTariffs.DataSource = oDataTable
        gvTariffs.DataBind()
    End Sub
    
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

    Protected Sub btnAddTariff_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ClearFields()
        Call InitFromName("0")
        lblLegendCustomer.Visible = True
        lblLegendEditing.Visible = False
        ddlCustomer.Visible = True
        lblCustomerName.Visible = False
        Call PopulateCustomerDropdown()
        btnAddTariff.Visible = False
        btnSave.Enabled = False
        Call SetEnabled(False)
        ddlCustomer.Focus()
        Call ShowDetailPanel()
    End Sub

    Protected Sub ClearFields()
        tbA4Jiffy.Text = String.Empty
        tbA5Jiffy.Text = String.Empty
        tbA3Jiffy.Text = String.Empty
        tbA4Box.Text = String.Empty
        tbA5Box.Text = String.Empty
        tbA3Box.Text = String.Empty
        tbSmallTube.Text = String.Empty
        tbMediumTube.Text = String.Empty
        tbLargeTube.Text = String.Empty
        tbFTC.Text = String.Empty
        tbInfill.Text = String.Empty
        tbOther.Text = String.Empty
        lblLegendEditing.Visible = False
        lblCustomerName.Text = String.Empty
        lblCustomerName.Visible = False
        ddlCustomer.Visible = False
    End Sub
    
    Protected Sub gvTariffs_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim tc As TableCell = gvr.Cells(1)
            Dim btn As Button
            If tc.Text = "AAAAAAAA" Then
                tc.Text = "<b>DEFAULT TARIFF</b>"
                btn = gvr.FindControl("btnRemove")
                btn.Visible = False
                btn = gvr.FindControl("btnEdit")
                btn.CommandArgument = "0"
            Else
                btn = gvr.FindControl("btnRemove")
                btn.CommandArgument = tc.Text.Split("~")(1).Trim
                btn = gvr.FindControl("btnEdit")
                btn.CommandArgument = tc.Text.Split("~")(1).Trim
            End If
        End If
    End Sub
    
    Protected Sub InitFromName(ByVal sCustomerName As String)
        Dim sSQL As String
        If sCustomerName = "0" Then
            sSQL = "SELECT * FROM ConsumablesCost WHERE CustomerKey = 0"
        Else
            sSQL = "SELECT * FROM ConsumablesCost cc INNER JOIN Customer c ON cc.CustomerKey = c.CustomerKey WHERE c.CustomerAccountCode = '" & sCustomerName.Replace("'", "''" & "'") & "'"
        End If
        Dim dr As DataRow = ExecuteQueryToDataTable(sSQL).Rows(0)
        tbA4Jiffy.Text = Format(dr("A4Jiffy"), "#,##0.00")
        tbA5Jiffy.Text = Format(dr("A5Jiffy"), "#,##0.00")
        tbA3Jiffy.Text = Format(dr("A3Jiffy"), "#,##0.00")
        tbA4Box.Text = Format(dr("A4Box"), "#,##0.00")
        tbA5Box.Text = Format(dr("A5Box"), "#,##0.00")
        tbA3Box.Text = Format(dr("A3Box"), "#,##0.00")
        tbSmallTube.Text = Format(dr("SmallTube"), "#,##0.00")
        tbMediumTube.Text = Format(dr("MedTube"), "#,##0.00")
        tbLargeTube.Text = Format(dr("LargeTube"), "#,##0.00")
        tbFTC.Text = Format(dr("FTC"), "#,##0.00")
        tbInfill.Text = Format(dr("Infill"), "#,##0.00")
        tbOther.Text = Format(dr("Other"), "#,##0.00")
    End Sub
    
    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Call ClearFields()
        lblLegendEditing.Visible = True
        lblCustomerName.Text = b.CommandArgument
        If lblCustomerName.Text = "0" Then
            lblCustomerName.Text = "DEFAULT TARIFF"
        End If
        lblCustomerName.Visible = True
        lblLegendCustomer.Visible = False
        ddlCustomer.Visible = False
        Call InitFromName(b.CommandArgument)
        btnAddTariff.Visible = False
        Call ShowDetailPanel()
    End Sub

    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        Dim sCustomerName As String = b.CommandArgument
        Dim sSQL As String = "DELETE FROM ConsumablesCost WHERE CustomerKey = (SELECT CustomerKey FROM Customer WHERE CustomerAccountCode = '" & sCustomerName.Replace("'", "''") & "') SELECT @@ROWCOUNT"
        If ExecuteQueryToDataTable(sSQL).Rows(0).Item(0) <> 1 Then
            WebMsgBox.Show("Unexpected row count returned from delete operation")
        End If
        Call BindTariffsGrid()
    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        btnAddTariff.Visible = True
        Call ShowCustomerPanel()
    End Sub

    Protected Function GetCustomerKeyFromName(ByVal sCustomerName As String) As Integer
        Dim sSQL As String
        If sCustomerName = "DEFAULT TARIFF" Then
            GetCustomerKeyFromName = 0
        Else
            sSQL = "SELECT CustomerKey FROM Customer WHERE CustomerAccountCode = '" & sCustomerName.Replace("'", "''") & "'"
            GetCustomerKeyFromName = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        End If
    End Function
    
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sSQL As String
        If lblLegendEditing.Visible Then
            sSQL = "UPDATE ConsumablesCost SET A4Jiffy = " & tbA4Jiffy.Text & ", A5Jiffy = " & tbA5Jiffy.Text & ", A3Jiffy = " & tbA3Jiffy.Text & ", A4Box = " & tbA4Box.Text & ", A5Box = " & tbA5Box.Text & ", A3Box = " & tbA3Box.Text & ", SmallTube = " & tbSmallTube.Text & ", MedTube = " & tbMediumTube.Text & ", LargeTube = " & tbLargeTube.Text & ", FTC = " & tbFTC.Text & ", Infill = " & tbInfill.Text & ", Other = " & tbOther.Text & ", LastUpdatedBy = " & Session("UserKey") & ", LastUpdatedOn = GETDATE() WHERE CustomerKey = " & GetCustomerKeyFromName(lblCustomerName.Text)
        Else
            sSQL = "INSERT INTO ConsumablesCost (CustomerKey, A4Jiffy, A5Jiffy, A3Jiffy,A4Box, A5Box, A3Box, SmallTube, MedTube, LargeTube, FTC, Infill, Other, LastUpdatedBy, LastUpdatedOn) VALUES (" & GetCustomerKeyFromName(ddlCustomer.SelectedValue) & ", " & tbA4Jiffy.Text & ", " & tbA5Jiffy.Text & ", " & tbA3Jiffy.Text & ", " & tbA4Box.Text & ", " & tbA5Box.Text & ", " & tbA3Box.Text & ", " & tbSmallTube.Text & ", " & tbMediumTube.Text & ", " & tbLargeTube.Text & ", " & tbFTC.Text & ", " & tbInfill.Text & ", " & tbOther.Text & ", " & Session("UserKey") & ", GETDATE())"
        End If
        Call ExecuteNonQuery(sSQL)
        Call BindTariffsGrid()
        btnAddTariff.Visible = True
        Call ShowCustomerPanel()
    End Sub

    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            btnSave.Enabled = True
            Call SetEnabled(True)
        Else
            btnSave.Enabled = False
            Call SetEnabled(False)
        End If
    End Sub

    Protected Sub SetEnabled(ByVal bEnabled As Boolean)
        tbA4Jiffy.Enabled = bEnabled
        tbA5Jiffy.Enabled = bEnabled
        tbA3Jiffy.Enabled = bEnabled
        tbA4Box.Enabled = bEnabled
        tbA5Box.Enabled = bEnabled
        tbA3Box.Enabled = bEnabled
        tbSmallTube.Enabled = bEnabled
        tbMediumTube.Enabled = bEnabled
        tbLargeTube.Enabled = bEnabled
        tbFTC.Enabled = bEnabled
        tbInfill.Enabled = bEnabled
        tbOther.Enabled = bEnabled
    End Sub
    
    Protected Sub btnConsumablesUsageReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not (IsDate(tbFromDate.Text) And IsDate(tbToDate.Text)) Then
            WebMsgBox.Show("Please ensure the FROM and TO dates are correct.")
            Exit Sub
        End If
        
        Dim sbSQL As New StringBuilder
        tbResult.Text = String.Empty
        sbSQL.Append("")
        sbSQL.Append("DECLARE @CustomerKey int, @CustomerName varchar(50), @StartDate smalldatetime, @EndDate smalldatetime, @TotalConsignmentCount int, @TotalConsumablesCost money ")
        sbSQL.Append("SET @StartDate = '" & tbFromDate.Text & " 00:01' ")
        sbSQL.Append("SET @EndDate = '" & tbToDate.Text & " 23:59' ")
        sbSQL.Append("PRINT 'CONSUMABLES USAGE REPORT' ")
        sbSQL.Append("PRINT 'FROM: ' + CAST(@StartDate AS varchar(20)) + ' TO ' + CAST(@EndDate AS varchar(20)) ")
        sbSQL.Append("PRINT '' ")
        sbSQL.Append("DECLARE c CURSOR FOR ")
        sbSQL.Append("SELECT CustomerKey, CustomerAccountCode ")
        sbSQL.Append("FROM Customer ")
        sbSQL.Append("WHERE ISNULL(AccountHandlerKey,0) > 0 AND CustomerStatusId = 'ACTIVE' ")
        sbSQL.Append("AND NOT (CustomerAccountCode LIKE 'demo%' OR CustomerAccountCode LIKE 'internal') ")
        sbSQL.Append("ORDER BY CustomerAccountCode ")
        sbSQL.Append("OPEN c ")
        sbSQL.Append("FETCH NEXT FROM c INTO @CustomerKey, @CustomerName ")
        sbSQL.Append("WHILE (@@FETCH_STATUS) = 0 ")
        sbSQL.Append("BEGIN ")
        sbSQL.Append("  SET @TotalConsumablesCost = (SELECT SUM(TotalCost) FROM ConsumablesUsed WHERE ConsignmentKey IN (SELECT [key] FROM Consignment WHERE CreatedOn BETWEEN @StartDate AND @EndDate AND CustomerKey = @CustomerKey)) ")
        sbSQL.Append("  IF ISNULL(@TotalConsumablesCost,0) > 0 ")
        sbSQL.Append("  BEGIN ")
        sbSQL.Append("    SET @TotalConsignmentCount = (SELECT COUNT(*) FROM Consignment WHERE CreatedOn BETWEEN @StartDate AND @EndDate AND CustomerKey = @CustomerKey) ")
        sbSQL.Append("    PRINT 'CUSTOMER: ' + @CustomerName ")
        sbSQL.Append("    PRINT 'Total consignments in period: ' + CAST(@TotalConsignmentCount AS varchar(10)) ")
        sbSQL.Append("    PRINT 'Total cost of consumables used: £' + CAST(@TotalConsumablesCost AS varchar(10)) ")
        sbSQL.Append("    PRINT '' ")
        sbSQL.Append("  END ")
        sbSQL.Append("  FETCH NEXT FROM c INTO @CustomerKey, @CustomerName ")
        sbSQL.Append("END ")
        sbSQL.Append("CLOSE c DEALLOCATE c ")
        sbSQL.Append("PRINT '[end of report]' ")
        sbSQL.Append("")
        sbSQL.Append("")
        sbSQL.Append("")
        Dim sSQL As String = sbSQL.ToString
        Call ExecuteQueryToDataTablePlus(sSQL, tbResult)
    End Sub
    
    Protected Sub btnOverlengthPicksReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not (IsDate(tbFromDate.Text) And IsDate(tbToDate.Text)) Then
            WebMsgBox.Show("Please ensure the FROM and TO dates are correct.")
            Exit Sub
        End If
        If Not IsNumeric(tbMaxNormalPickDuration.Text) Then
            WebMsgBox.Show("Please indicate the maximum number of minutes, eg 5, for a pick.")
            Exit Sub
        End If
        
        Dim sbSQL As New StringBuilder
        tbResult.Text = String.Empty
        sbSQL.Append("")
        sbSQL.Append("DECLARE @CustomerKey int, @CustomerName varchar(50), @StartDate smalldatetime, @EndDate smalldatetime, @PickMinsTotal int, @PickMinsThreshold int, @TotalConsignmentCount int, @ExcessPickMinsConsignmentCount int ")
        sbSQL.Append("SET @PickMinsThreshold = " & tbMaxNormalPickDuration.Text & " ")
        sbSQL.Append("SET @StartDate = '" & tbFromDate.Text & " 00:01' ")
        sbSQL.Append("SET @EndDate = '" & tbToDate.Text & " 23:59' ")
        sbSQL.Append("PRINT 'OVERLENGTH PICK MINUTES REPORT' ")
        sbSQL.Append("PRINT 'FROM: ' + CAST(@StartDate AS varchar(20)) + ' TO ' + CAST(@EndDate AS varchar(20)) ")
        sbSQL.Append("PRINT 'Reporting picks longer than " & tbMaxNormalPickDuration.Text & " minutes'")
        sbSQL.Append("PRINT '' ")
        sbSQL.Append("DECLARE c CURSOR FOR ")
        sbSQL.Append("SELECT CustomerKey, CustomerAccountCode ")
        sbSQL.Append("FROM Customer ")
        sbSQL.Append("WHERE ISNULL(AccountHandlerKey,0) > 0 AND CustomerStatusId = 'ACTIVE' ")
        sbSQL.Append("AND NOT (CustomerAccountCode LIKE 'demo%' OR CustomerAccountCode LIKE 'internal') ")
        sbSQL.Append("ORDER BY CustomerAccountCode ")
        sbSQL.Append("OPEN c ")
        sbSQL.Append("FETCH NEXT FROM c INTO @CustomerKey, @CustomerName ")
        sbSQL.Append("WHILE (@@FETCH_STATUS) = 0 ")
        sbSQL.Append("BEGIN ")
        sbSQL.Append("  SET @TotalConsignmentCount = (SELECT COUNT(*) FROM Consignment WHERE CreatedOn BETWEEN @StartDate AND @EndDate AND CustomerKey = @CustomerKey) ")
        sbSQL.Append("  SET @PickMinsTotal = (SELECT SUM(ISNULL(PickMins,0)) FROM Consignment WHERE CreatedOn BETWEEN @StartDate AND @EndDate AND ISNULL(PickMins,0) > @PickMinsThreshold AND CustomerKey = @CustomerKey) ")
        sbSQL.Append("  IF @PickMinsTotal IS NULL PRINT '....... no overlength picks for customer ' + @CustomerName ")
        sbSQL.Append("  ELSE BEGIN ")
        sbSQL.Append("    PRINT '' ")
        sbSQL.Append("    PRINT 'CUSTOMER: ' + @customername ")
        sbSQL.Append("    PRINT 'Total consignments in period: ' + CAST(@TotalConsignmentCount AS varchar(10)) ")
        sbSQL.Append("    PRINT 'Total pick mins: ' + CAST(@PickMinsTotal AS varchar(10)) ")
        sbSQL.Append("    SET @ExcessPickMinsConsignmentCount = (SELECT COUNT(*) FROM Consignment WHERE CreatedOn BETWEEN @StartDate AND @EndDate AND CustomerKey = @CustomerKey AND ISNULL(PickMins,0) > @PickMinsThreshold) ")
        sbSQL.Append("    PRINT 'Consignments with excess pick mins: ' + CAST(@ExcessPickMinsConsignmentCount AS varchar(10)) ")
        sbSQL.Append("    SET @PickMinsTotal = @PickMinsTotal - (@PickMinsThreshold * @ExcessPickMinsConsignmentCount) ")
        sbSQL.Append("    PRINT 'Total excess pick mins: ' + CAST(@PickMinsTotal AS varchar(10)) ")
        sbSQL.Append("    PRINT '' ")
        sbSQL.Append("  END")
        sbSQL.Append("  FETCH NEXT FROM c INTO @CustomerKey, @CustomerName ")
        sbSQL.Append("END ")
        sbSQL.Append("CLOSE c DEALLOCATE c ")
        sbSQL.Append("PRINT '' ")
        sbSQL.Append("PRINT '[end of report]' ")
        Dim sSQL As String = sbSQL.ToString
        Call ExecuteQueryToDataTablePlus(sSQL, tbResult)
    End Sub

    Protected Function ExecuteQueryToDataTablePlus(ByVal sQuery As String, Optional ByVal oControl As Control = Nothing) As DataTable
        Dim oDataTable As New DataTable
        Dim oConn As New SqlConnection(gsConn)
        Dim oAdapter As New SqlDataAdapter(sQuery, oConn)
        Dim oCmd As SqlCommand = New SqlCommand(sQuery, oConn)
        Try
            oConn.Open()
            goControl = oControl
            AddHandler oConn.InfoMessage, Function(sender, f) ExecuteQueryToDataTablePlusAnonymousMethod(sender, f)
            oAdapter.Fill(oDataTable)
        Catch ex As Exception
            WebMsgBox.Show("Error in ExecuteQueryToDataTablePlus executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTablePlus = oDataTable
    End Function

    Private Function ExecuteQueryToDataTablePlusAnonymousMethod(ByVal sender As Object, ByVal f As SqlInfoMessageEventArgs) As Boolean
        Dim tb As TextBox = Nothing
        Dim lbl As Label = Nothing
        Dim lb As ListBox = Nothing
        Dim ddl As DropDownList = Nothing
        If TypeOf (goControl) Is TextBox Then
            tb = goControl
            tb.Text += Constants.vbLf + f.Message
        ElseIf TypeOf goControl Is Label Then
            lbl = goControl
            lbl.Text += "<br />" + f.Message
        ElseIf TypeOf goControl Is ListBox Then
            lb = goControl
            lb.Items.Add(f.Message)
        ElseIf TypeOf goControl Is DropDownList Then
            ddl = goControl
            ddl.Items.Add(f.Message)
        End If
        Return True
    End Function

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>File Upload</title>
    <style type="text/css">

.light {
    color: silver
}    

.informational {
    font-size:xx-small
}
    </style>
    </head>
<body>
    <form id="form1" runat="server">
    <div>
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="left" style="width: 35%">
        &nbsp;
        <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Consumables Tariffs & Usage"/>
        </td>
                <td style="width: 55%">
                    <asp:Button ID="btnAddTariff" runat="server" Text="add customer-specific tariff" Width="200px" onclick="btnAddTariff_Click" />
                &nbsp;</td>
                <td style="width: 10%">
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlCustomer" runat="server" Width="100%">
            <strong>
                &nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br />
            <table style="width: 100%">
                <tr>
                    <td style="width: 100%">
                        <asp:GridView ID="gvTariffs" runat="server" CellPadding="2" 
                            Font-Names="Verdana" Font-Size="XX-Small" Width="95%" 
                            EmptyDataText="no entries found" OnRowDataBound="gvTariffs_RowDataBound">
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnEdit" runat="server" Text="edit" OnClick="btnEdit_Click" />
                                        <asp:Button ID="btnRemove" runat="server" Text="remove" onclick="btnRemove_Click"  />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                            <PagerStyle HorizontalAlign="Center" />
                            <EmptyDataTemplate>
                                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="no tariffs found"/>
                            </EmptyDataTemplate>    
                        </asp:GridView>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlDetail" runat="server" Visible="false" Width="100%" Font-Names="Verdana">
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td style="width: 20%">
                </td>
                <td style="width: 60%">
                </td>
                <td style="width: 20%">
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendCustomer" runat="server" Font-Size="Small">Customer:</asp:Label>
                    <asp:Label ID="lblLegendEditing" runat="server" Font-Bold="True" Font-Size="Small" Text="Editing"/>
                </td>
                <td valign="top">
                    <asp:DropDownList ID="ddlCustomer" runat="server" Font-Names="Verdana" Font-Size="XX-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged" />
                    &nbsp;<asp:Label ID="lblCustomerName" runat="server" Font-Size="Small"></asp:Label>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right" style="width: 20%">
                    A4 Jiffy:
                </td>
                <td>
                    <asp:TextBox ID="tbA4Jiffy" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA4Jiffy" runat="server" ControlToValidate="tbA4Jiffy" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA4Jiffy" runat="server" ControlToValidate="tbA4Jiffy" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right" style="width: 20%">
                    A5 Jiffy:</td>
                <td style="width: 60%">
                    <asp:TextBox ID="tbA5Jiffy" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA5Jiffy" runat="server" ControlToValidate="tbA5Jiffy" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA5Jiffy" runat="server" ControlToValidate="tbA5Jiffy" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    A3 Jiffy:
                </td>
                <td style="width: 60%">
                    <asp:TextBox ID="tbA3Jiffy" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA3Jiffy" runat="server" ControlToValidate="tbA3Jiffy" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA3Jiffy" runat="server" ControlToValidate="tbA3Jiffy" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    A4 Box:
                </td>
                <td>
                    <asp:TextBox ID="tbA4Box" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA4Box" runat="server" ControlToValidate="tbA4Box" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA4Box" runat="server" ControlToValidate="tbA4Box" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    A5 Box:</td>
                <td>
                    <asp:TextBox ID="tbA5Box" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA5Box" runat="server" ControlToValidate="tbA5Box" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA5Box" runat="server" ControlToValidate="tbA5Box" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    A3 Box:
                </td>
                <td>
                    <asp:TextBox ID="tbA3Box" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvA3Box" runat="server" ControlToValidate="tbA3Box" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvA3Box" runat="server" ControlToValidate="tbA3Box" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    Small tube:
                </td>
                <td>
                    <asp:TextBox ID="tbSmallTube" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvSmallTube" runat="server" ControlToValidate="tbSmallTube" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvSmallTube" runat="server" ControlToValidate="tbSmallTube" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    Medium tube:
                </td>
                <td style="width: 60%">
                    <asp:TextBox ID="tbMediumTube" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvMediumTube" runat="server" ControlToValidate="tbMediumTube" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvMediumTube" runat="server" ControlToValidate="tbMediumTube" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    Large tube:
                </td>
                <td style="width: 60%">
                    <asp:TextBox ID="tbLargeTube" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvLargeTube" runat="server" ControlToValidate="tbLargeTube" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvLargeTube" runat="server" ControlToValidate="tbLargeTube" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    FTC:
                </td>
                <td>
                    <asp:TextBox ID="tbFTC" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvFTC" runat="server" ControlToValidate="tbFTC" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvFTC" runat="server" ControlToValidate="tbFTC" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    Infill:
                </td>
                <td>
                    <asp:TextBox ID="tbInfill" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvInfill" runat="server" ControlToValidate="tbInfill" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvInfill" runat="server" ControlToValidate="tbInfill" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                    Other:
                </td>
                <td>
                    <asp:TextBox ID="tbOther" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="120px" MaxLength="6"/>
                    &nbsp;
                    <asp:RequiredFieldValidator ID="rfvOther" runat="server" ControlToValidate="tbOther" ErrorMessage="required!"/>
                    <asp:RangeValidator ID="rvOther" runat="server" ControlToValidate="tbOther" ErrorMessage="must be numeric!" MaximumValue="100" MinimumValue="0" Type="Currency"/>
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                    &nbsp; &nbsp;&nbsp;
                </td>
                <td/>
            </tr>
            <tr>
                <td align="right">
                </td>
                <td>
                    <asp:Button ID="btnSave" runat="server" Text="save" Width="100px" OnClick="btnSave_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" Text="cancel" Width="100px" CausesValidation="False" onclick="btnCancel_Click" />
                </td>
                <td/>
            </tr>
        </table>
        </asp:Panel>
        <asp:Panel ID="pnlHelp" runat="server" Visible="true" Width="100%" Font-Names="Verdana">
            <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
                <tr>
                    <td style="width: 20%">
                    </td>
                    <td style="width: 60%">
                    </td>
                    <td style="width: 20%">
                    </td>
                </tr>
                <tr>
                    <td />
                    <td>
                        NOTE: The <b>DEFAULT TARIFF</b> applies to all customer accounts that do not have
                        a customer-specific tariff defined.
                    </td>
                    <td />
                </tr>
                <tr>
                    <td />
                    &nbsp;<td>
                        &nbsp;
                    </td>
                    <td />
                    &nbsp;</tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="pnlUsage" runat="server" Visible="true" Width="100%" Font-Names="Verdana">
            <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
                <tr>
                    <td style="width: 0%"/>
                    <td style="width: 1000%">
                        <asp:Label ID="lblTitle0" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small"
                            Text="Consumables Usage &amp; Overlength Pick Duration Reports" />
                    </td>
                    <td style="width: 5%"/>
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                        From:
                        <asp:TextBox ID="tbFromDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="12" Width="80px" />
                        &nbsp;<a id="imgCalendarButton1" runat="server" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')"
                            visible="true"><img id="Img1" runat="server" alt="" border="0" ie:visible="true"
                                src="images/SmallCalendar.gif" visible="false" /></a> <span id="spnDateExample1"
                                    runat="server" class="informational light" style="white-space: nowrap" visible="true">
                                    (eg&nbsp;1-Jan-2010)</span> To:
                        <asp:TextBox ID="tbToDate" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="12" Width="80px" />
                        &nbsp;<a id="imgCalendarButton2" runat="server" href="javascript:;" onclick="window.open('../PopupCalendar4.aspx?textbox=tbFromDate','cal','width=300,height=305,left=270,top=180')"
                            visible="true"><img id="Img2" runat="server" alt="" border="0" ie:visible="true"
                                src="images/SmallCalendar.gif" visible="false" /></a> <span id="spnDateExample2"
                                    runat="server" class="informational light" style="white-space: nowrap" visible="true">
                                    (eg&nbsp;31-Jan-2010)</span> &nbsp; &nbsp;Overlength picks are picks
                        longer than
                        <asp:TextBox ID="tbMaxNormalPickDuration" runat="server" Font-Names="Verdana" Font-Size="XX-Small"
                            MaxLength="6" Width="30px">5</asp:TextBox>
                        &nbsp;mins
                    </td>
                </tr>
                <tr>
                    <td />
                    <td />
                    <td />
                </tr>
                <tr>
                    <td />
                    <td colspan="2">
                        <asp:Button ID="btnConsumablesUsageReport" runat="server" OnClick="btnConsumablesUsageReport_Click"
                            Text="consumables usage report" Width="270px" />
                        &nbsp;<asp:Button ID="btnOverlengthPicksReport" runat="server" OnClick="btnOverlengthPicksReport_Click"
                            Text="overlength picks report" Width="270px" />
                    </td>
                </tr>
                <tr>
                    <td />
                    <td>
                        <asp:TextBox ID="tbResult" runat="server" Rows="20" TextMode="MultiLine" 
                            Width="95%" />
                    </td>
                    <td />
                </tr>
            </table>
        </asp:Panel>
        </div>
    </form>
</body>
</html>
