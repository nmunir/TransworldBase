<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.SqlTypes" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections.Generic" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    ' CHANGES
    ' one level of UNDO
    ' all columns sortable
    ' sort toggles up-down, down-up
    

    Dim gsConn As String = ConfigurationManager.AppSettings("AIMSRootConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call InitCustomerDropdown()
            Call InitPeriodDropdown()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Job Pricing"
    End Sub

    Protected Sub InitCustomerDropdown()
        Dim sSQL As String
        'sSQL = "SELECT DISTINCT CustomerAccountCode, CustomerKey FROM Customer WHERE WarehouseID = 1 AND DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        sSQL = "SELECT DISTINCT CustomerAccountCode, CustomerKey FROM Customer WHERE DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Clear()
        ddlCustomer.Items.Add(New ListItem("- please select -", "0"))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
    End Sub

    Protected Sub InitPeriodDropdown()
        ddlPeriod.Items.Clear()
        ddlPeriod.Items.Add(New ListItem("- please select -", "0"))
        For i As Int32 = -3 To 3
            Dim dtStartDate As Date = DateAdd(DateInterval.Month, i, Date.Now)
            Dim sDate As String = Format(dtStartDate, "MMM") & " " & Format(dtStartDate, "yyyy")
            Dim dtNextDate As Date = DateAdd(DateInterval.Month, i + 1, Date.Now)
            Dim sDateClause As String = "CreatedOn >= '1-" & Format(dtStartDate, "MMM") & "-" & Format(dtStartDate, "yyyy") & " 00:01' AND CreatedOn < '1-" & Format(dtNextDate, "MMM") & "-" & Format(dtNextDate, "yyyy") & "'"
            ddlPeriod.Items.Add(New ListItem(sDate, sDateClause))
        Next
    End Sub

    Protected Sub ddlCustomer_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        If ddlPeriod.SelectedIndex > 0 Then
            Call PopulateOrders()
            If ddlPeriod.SelectedValue = "0" Then
                'Call DisplayOrderCount()
                lblOrderCount.Text = String.Empty
            End If
        End If
    End Sub

    Protected Sub ddlPeriod_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        If ddlCustomer.SelectedIndex > 0 Then
            Call PopulateOrders()
            If ddlCustomer.SelectedValue = "0" Then
                'Call DisplayOrderCount()
                lblOrderCount.Text = String.Empty
            End If
        End If
    End Sub
    
    Protected Sub DisplayOrderCount()
        lblOrderCount.Text = "- <b>" & gvOrders.Rows.Count & "</b> order(s) found"
    End Sub
    
    Protected Sub PopulateOrders(Optional ByVal sSortExpresion As String = "[key]", Optional ByVal sSortOrder As String = " DESC", Optional ByVal bExport As Boolean = False)
        'Dim sSQL As String = "SELECT [key], CashOnDelAmount, REPLACE(CONVERT(VARCHAR(9), CreatedOn, 6), ' ', '-') 'CreatedOn', AWB, StateID, [dbo].JobPricingCountLines([key]) 'Lines', [dbo].JobPricingCountUnits([key]) 'Units', NOP, Weight, ISNULL(CAST(ConsignmentTransport AS varchar(10)),'') 'Method', CustomerRef1, CustomerRef2, SpecialInstructions, CneeName , CneeAddr1 , CneeTown , CneePostCode FROM Consignment WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND " & ddlPeriod.SelectedValue
        If ddlPeriod.SelectedValue = "0" Then
            gvOrders.DataSource = Nothing
            gvOrders.DataBind()
            Call DisplayOrderCount()
            Exit Sub
        End If
        
        Dim sSQL As String = "SELECT [key], CashOnDelAmount, REPLACE(CONVERT(VARCHAR(9), CreatedOn, 6), ' ', '-') 'CreatedOn', AWB, StateID, [dbo].JobPricingCountLines([key]) 'Lines', [dbo].JobPricingCountUnits([key]) 'Units', NOP, Weight, CASE WHEN ConsignmentTransport IS NULL  THEN '' WHEN ConsignmentTransport = 0 THEN 'C' WHEN ConsignmentTransport = 1 THEN 'M' END 'Method', CustomerRef1, CustomerRef2, SpecialInstructions, CneeName , CneeAddr1 , CneeTown , CneePostCode FROM Consignment WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND " & ddlPeriod.SelectedValue
        ' 
        If cbCompletedOrdersOnly.Checked Then
            sSQL += " AND StateId = 'WITH_OPERATIONS' "
        End If
        If rbAllOrders.Checked Then
            'sSQL = "SELECT CashOnDelAmount, REPLACE(CONVERT(VARCHAR(9), CreatedOn, 6), ' ', '-') 'CreatedOn', AWB, NOP, Weight, CustomerRef1, CustomerRef2, SpecialInstructions, CneeName , CneeAddr1 , CneeTown , CneePostCode FROM Consignment WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND " & ddlPeriod.SelectedValue
            'ElseIf rbCompletedOrders.Checked Then
            '    sSQL = sSQL & " AND StateId = 'WITH_OPERATIONS'"
        ElseIf rbUnpricedOrders.Checked Then
            sSQL = sSQL & " AND ISNULL(CashOnDelAmount,0) = 0"
        Else
            Dim sMatch As String = "'%" & tbMatch.Text.Replace("'", "''") & "%'"
            Dim sMatchClause As String = " AND (AWB LIKE " & sMatch & " OR CustomerRef1 LIKE " & sMatch & " OR CustomerRef2 LIKE " & sMatch & " OR SpecialInstructions LIKE " & sMatch & " OR CneeName LIKE " & sMatch & " OR CneeAddr1 Like " & sMatch & " OR CneeTown LIKE " & sMatch & " OR CneePostCode LIKE " & sMatch & ")"
            sSQL = sSQL & sMatchClause
        End If
        sSQL = sSQL & " ORDER BY " & sSortExpresion & " " & sSortOrder
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If bExport = False Then
            gvOrders.DataSource = dt
            gvOrders.DataBind()
            If dt.Rows.Count > 0 Then
                lnkbtnExportToExcel.Visible = True
            Else
                lnkbtnExportToExcel.Visible = False
            End If
        Else
            Response.Clear()
            Response.ContentType = "text/csv"
            Dim sResponseValue As New StringBuilder
            sResponseValue.Append("attachment; filename=""")
            sResponseValue.Append(ddlCustomer.SelectedItem.Text)
            sResponseValue.Append("_")
            sResponseValue.Append(ddlPeriod.SelectedItem.Text)
            sResponseValue.Append("_")
            sResponseValue.Append(Format(Date.Now, "ddMMMyyyy_hhmmss"))
            sResponseValue.Append(".csv")
            sResponseValue.Append("""")
            Response.AddHeader("Content-Disposition", sResponseValue.ToString)

            Dim sItem As String
            For Each dr As DataRow In dt.Rows
                For i = 0 To dt.Columns.Count - 1
                    sItem = dr(i) & String.Empty
                    sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                    sItem = ControlChars.Quote & sItem & ControlChars.Quote
                    Response.Write(sItem)
                    Response.Write(",")
                Next
                Response.Write(vbCrLf)
            Next
            Response.End()
        End If
        Call DisplayOrderCount()
    End Sub
    
    Protected Sub cbMasterSelect_CheckedChanged(sender As Object, e As System.EventArgs)
        Dim cbMasterSelect As CheckBox = sender
        For Each gvr As GridViewRow In gvOrders.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim thisgvr As GridViewRow = gvr
                Dim cb As CheckBox
                cb = thisgvr.Cells(0).Controls(1)
                cb.Checked = cbMasterSelect.Checked
            End If
        Next
    End Sub
    
    Protected Sub rbAllOrders_CheckedChanged(sender As Object, e As System.EventArgs)
        Call PopulateOrders()
        tbMatch.Text = String.Empty
    End Sub

    Protected Sub rbUnpricedOrders_CheckedChanged(sender As Object, e As System.EventArgs)
        Call PopulateOrders()
        tbMatch.Text = String.Empty
    End Sub

    Protected Sub SavePricesForUndo()
        Dim sSQL As String
        sSQL = "DELETE FROM JobPricingUndo WHERE UserKey = " & Session("UserKey") & " AND CustomerKey = " & ddlCustomer.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
        sSQL = "INSERT INTO JobPricingUndo (ConsignmentKey, Price, UserKey, CustomerKey) "
        sSQL += "SELECT [key] 'ConsignmentKey', CashOnDelAmount 'Price', " & Session("UserKey") & ", " & ddlCustomer.SelectedValue & " FROM Consignment WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND " & ddlPeriod.SelectedValue
        Call ExecuteQueryToDataTable(sSQL)
        lnkbtnUndo.Enabled = True
    End Sub
    
    Protected Sub btnApplyPricingToCheckedOrders_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call ApplyPricingToCheckedOrders()
    End Sub
    
    Protected Sub ApplyPricingToCheckedOrders()
        tbPricePerOrder.Text = tbPricePerOrder.Text.Trim
        tbPricePerUnit.Text = tbPricePerUnit.Text.Trim
        tbPriceFirstLine.Text = tbPriceFirstLine.Text.Trim
        tbPriceMoreLines.Text = tbPriceMoreLines.Text.Trim
        Dim bPriceFound As Boolean = False
        
        If tbPricePerOrder.Text <> String.Empty Then
            If Not IsNumeric(tbPricePerOrder.Text) Then
                tbPricePerOrder.Focus()
                WebMsgBox.Show("PER ORDER PRICE: please enter a valid price, eg 6.80")
                Exit Sub
            End If
            bPriceFound = True
        End If
        
        If tbPricePerUnit.Text <> String.Empty Then
            If Not IsNumeric(tbPricePerUnit.Text) Then
                tbPricePerUnit.Focus()
                WebMsgBox.Show("PER UNIT PRICE: please enter a valid price, eg 6.80")
                Exit Sub
            End If
            bPriceFound = True
        End If
        
        If tbPriceFirstLine.Text <> String.Empty Then
            If Not IsNumeric(tbPriceFirstLine.Text) Then
                tbPriceFirstLine.Focus()
                WebMsgBox.Show("FIRST LINE PRICE: please enter a valid price, eg 6.80")
                Exit Sub
            End If
            bPriceFound = True
        End If
        
        If tbPriceMoreLines.Text <> String.Empty Then
            If Not IsNumeric(tbPriceMoreLines.Text) Then
                tbPriceMoreLines.Focus()
                WebMsgBox.Show("MORE LINES PRICE: please enter a valid price, eg 6.80")
                Exit Sub
            End If
            bPriceFound = True
        End If
        
        'If tbPricePerOrder.Text = String.Empty Then
        '    Exit Sub
        'End If
        'If Not IsNumeric(tbPricePerOrder.Text) Then
        '    WebMsgBox.Show("Please enter a valid price, eg 6.80")
        '    Exit Sub
        'End If
        
        If bPriceFound Then
            Call SavePricesForUndo()
            For Each gvr As GridViewRow In gvOrders.Rows
                If gvr.RowType = DataControlRowType.DataRow Then
                    Dim cb As CheckBox
                    cb = gvr.Cells(0).Controls(1)
                    If cb.Checked Then
                        Dim dblPrice As Double = 0.0
                        Dim sAWB As String = gvr.Cells(3).Text
                        If IsNumeric(tbPricePerOrder.Text) Then
                            dblPrice += CDbl(tbPricePerOrder.Text)
                        End If
                        If IsNumeric(tbPricePerUnit.Text) Then
                            Dim nUnitCount As Int32 = ExecuteQueryToDataTable("SELECT dbo.JobPricingCountUnits(" & sAWB & ")").Rows(0).Item(0)
                            dblPrice += nUnitCount * CInt(tbPricePerUnit.Text)
                        End If
                        If IsNumeric(tbPriceFirstLine.Text) Or IsNumeric(tbPriceMoreLines.Text) Then
                            Dim nLineCount As Int32 = ExecuteQueryToDataTable("SELECT dbo.JobPricingCountLines(" & sAWB & ")").Rows(0).Item(0)
                            If IsNumeric(tbPriceFirstLine.Text) Then
                                dblPrice += CDbl(tbPriceFirstLine.Text)
                            End If
                            If IsNumeric(tbPriceMoreLines.Text) Then
                                dblPrice += CDbl(tbPriceMoreLines.Text * (nLineCount - 1))
                            End If
                        End If
                        Dim sSQL As String = "UPDATE Consignment SET CashOnDelAmount = " & dblPrice.ToString & " WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND [key] = " & sAWB
                        Call ExecuteQueryToDataTable(sSQL)
                    End If
                End If
            Next
            'rbCompletedOrders.Checked = True
            'rbAllOrders.Checked = True
            Call PopulateOrders()
        End If
    End Sub
    
    Protected Sub btnGoSearch_Click(sender As Object, e As System.EventArgs)
        'rbCompletedOrders.Checked = False
        rbAllOrders.Checked = False
        rbUnpricedOrders.Checked = False
        rbOrdersMatching.Checked = True
        Call PopulateOrders()
    End Sub
    
    Protected Sub gvOrders_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Dim gvr As GridViewRow = e.Row
        gvr.Cells(0).BackColor = Drawing.Color.White
        If gvr.RowType = DataControlRowType.DataRow Then
            Dim hid As HiddenField = gvr.Cells(0).FindControl("hidStateID")
            If hid.Value = "WITH_OPERATIONS" Then
                gvr.Cells(0).BackColor = Drawing.Color.Green
            ElseIf hid.Value = "CANCELLED" Then
                gvr.Cells(0).BackColor = Drawing.Color.Red
            Else
                gvr.Cells(0).BackColor = Drawing.Color.Blue
            End If
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection: " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function

    Protected Sub cbCompletedOrdersOnly_CheckedChanged(sender As Object, e As System.EventArgs)
        Call PopulateOrders()
    End Sub
    
    Protected Sub rbOrdersMatching_CheckedChanged(sender As Object, e As System.EventArgs)
        tbMatch.Focus()
    End Sub
    
    Protected Sub lnkbtnExportToExcel_Click(sender As Object, e As System.EventArgs)
        Call PopulateOrders(bExport:=True)
    End Sub
    
    Protected Sub lnkbtnUndo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim sSQL As String
        sSQL = "UPDATE Consignment SET CashOnDelAmount = (SELECT Price FROM JobPricingUndo WHERE JobPricingUndo.ConsignmentKey = Consignment.[key]) "
        'sSQL += "WHERE Consignment.[key] IN (SELECT ConsignmentKey FROM JobPricingUndo WHERE UserKey = " & Session("UserKey") & " AND CustomerKey = " & ddlCustomer.SelectedValue & ")"
        sSQL += "WHERE Consignment.[key] IN (SELECT ConsignmentKey FROM JobPricingUndo WHERE UserKey = " & Session("UserKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
        lnkbtn.Enabled = False
        Call PopulateOrders()
    End Sub
    
    Protected Sub gvOrders_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        If e.SortExpression <> psSortExpression Then
            Call PopulateOrders(e.SortExpression, "ASC")
            psSortDirection = "ASC"
        Else
            If psSortDirection = "DESC" Then
                Call PopulateOrders(e.SortExpression, "ASC")
                psSortDirection = "ASC"
            Else
                Call PopulateOrders(e.SortExpression, "DESC")
                psSortDirection = "DESC"
            End If
        End If
        psSortExpression = e.SortExpression
    End Sub
    
    Property psSortExpression() As String
        Get
            Dim o As Object = ViewState("JP_SortExpression")
            If o Is Nothing Then
                Return "[key]"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("JP_SortExpression") = Value
        End Set
    End Property

    Property psSortDirection() As String
        Get
            Dim o As Object = ViewState("JP_SortDirection")
            If o Is Nothing Then
                Return "ASC"
            End If
            Return CStr(o)
        End Get
        Set(ByVal Value As String)
            ViewState("JP_SortDirection") = Value
        End Set
    End Property

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Job Pricing</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: xx-small; font-family: Verdana">
        <main:Header ID="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width: 50%; white-space: nowrap">
                </td>
                <td style="width: 50%; white-space: nowrap" align="right">
                </td>
            </tr>
        </table>
        <table width="95%">
            <tr>
                <td colspan="2">
                    <strong>
                        <asp:Label ID="lblTitle" runat="server" Font-Names="Verdana" Font-Size="Small">Job Pricing</asp:Label>
                    </strong>
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="top">
                    &nbsp;
                </td>
                <td style="width: 90%">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    Customer:
                </td>
                <td style="width: 90%">
                    <asp:DropDownList ID="ddlCustomer" runat="server" Font-Names="Verdana" Font-Size="Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged" />
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    Period:
                </td>
                <td style="width: 90%">
                    <asp:DropDownList ID="ddlPeriod" runat="server" Font-Names="Verdana" Font-Size="Small"
                        AutoPostBack="True" OnSelectedIndexChanged="ddlPeriod_SelectedIndexChanged" />
                &nbsp;<asp:Label ID="lblOrderCount" runat="server"/>
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    Filter:
                </td>
                <td style="width: 90%">
                    <asp:RadioButton ID="rbAllOrders" runat="server" AutoPostBack="True" 
                        oncheckedchanged="rbAllOrders_CheckedChanged" Text="all orders" 
                        Checked="True" GroupName="items" />
                    <asp:RadioButton ID="rbUnpricedOrders" runat="server" AutoPostBack="True" GroupName="items"
                        Text="unpriced orders" OnCheckedChanged="rbUnpricedOrders_CheckedChanged" />
                    <asp:RadioButton ID="rbOrdersMatching" runat="server" Text="orders matching:" 
                        GroupName="items" oncheckedchanged="rbOrdersMatching_CheckedChanged" />
                    &nbsp;<asp:TextBox ID="tbMatch" runat="server" Width="104px" Font-Names="Arial" 
                        Font-Size="XX-Small"></asp:TextBox>
                &nbsp;<asp:Button ID="btnGoSearch" runat="server" onclick="btnGoSearch_Click" 
                        Text="go" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="cbCompletedOrdersOnly" runat="server" Checked="True" 
                        Text="completed orders only" Font-Italic="True" AutoPostBack="True" 
                        oncheckedchanged="cbCompletedOrdersOnly_CheckedChanged" />
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    Per order(£):
                </td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbPricePerOrder" runat="server" Width="70px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    Per unit (£):</td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbPricePerUnit" runat="server" Width="70px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    First line (£):</td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbPriceFirstLine" runat="server" Width="70px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    More lines (£)</td>
                <td style="width: 90%" valign="middle">
                    <asp:TextBox ID="tbPriceMoreLines" runat="server" Width="70px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="middle">
                    &nbsp;</td>
                <td style="width: 90%" valign="middle">
                    <asp:Button ID="btnApplyPricingToCheckedOrders" runat="server" Text="Apply Pricing to Checked Orders" 
                        OnClick="btnApplyPricingToCheckedOrders_Click" />
                &nbsp;&nbsp;&nbsp;
                    <asp:LinkButton ID="lnkbtnUndo" runat="server" onclick="lnkbtnUndo_Click" Enabled="False">undo last price update</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td align="right" style="width: 10%" valign="top">
                    &nbsp;
                </td>
                <td style="width: 90%">
                    <asp:LinkButton ID="lnkbtnExportToExcel" runat="server" 
                        onclick="lnkbtnExportToExcel_Click">export order list to Excel</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td align="center" style="width: 10%" valign="top">
                    <br />
                    <br />
                    <br />
                    &nbsp;
                    colour key (first column):              colour key (first column):<br />
                    <br />
                    GREEN - completed<br />
                    RED - cancelled<br />
                    BLUE - in progress</td>
                <td style="width: 90%">
                    <asp:GridView ID="gvOrders" runat="server" Width="100%" CellPadding="2" AutoGenerateColumns="False" OnRowDataBound="gvOrders_RowDataBound" AllowSorting="True" OnSorting="gvOrders_Sorting">
                        <AlternatingRowStyle BackColor="#FFFFCC" />
                        <Columns>
                            <asp:TemplateField>
                                <HeaderTemplate>
                                    <asp:CheckBox ID="cbMasterSelect" runat="server" OnCheckedChanged="cbMasterSelect_CheckedChanged"
                                        AutoPostBack="True" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbSelect" runat="server" />
                                    <asp:HiddenField ID="hidStateID" runat="server"  Value='<%# DataBinder.Eval(Container.DataItem,"StateID") %>'/>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="CashOnDelAmount" HeaderText="Price" ReadOnly="True" SortExpression="CashOnDelAmount"
                                HtmlEncode="False" DataFormatString="{0:C2}" />
                            <asp:BoundField DataField="CreatedOn" HeaderText="Date" ReadOnly="True" 
                                SortExpression="CreatedOn" >
                            <ItemStyle Wrap="False" />
                            </asp:BoundField>
                            <asp:BoundField DataField="AWB" HeaderText="AWB" ReadOnly="True" SortExpression="AWB" />
                            <asp:BoundField DataField="Lines" HeaderText="Lines" ReadOnly="True" SortExpression="Lines" />
                            <asp:BoundField DataField="Units" HeaderText="Units" ReadOnly="True" SortExpression="Units" />
                            <asp:BoundField DataField="NOP" HeaderText="Pieces" ReadOnly="True" SortExpression="NOP" />
                            <asp:BoundField DataField="Weight" HeaderText="Weight" ReadOnly="True" SortExpression="Weight" />
                            <asp:BoundField DataField="Method" HeaderText="Method" ReadOnly="True" SortExpression="Method" />
                            <asp:BoundField DataField="CustomerRef1" HeaderText="Cust Ref 1" ReadOnly="True"
                                SortExpression="CustomerRef1" />
                            <asp:BoundField DataField="CustomerRef2" HeaderText="Cust Ref 2" ReadOnly="True"
                                SortExpression="CustomerRef2" />
                            <asp:BoundField DataField="SpecialInstructions" HeaderText="Special Instructions"
                                ReadOnly="True" SortExpression="SpecialInstructions" />
                            <asp:BoundField DataField="CneeName" HeaderText="Name" ReadOnly="True" SortExpression="CneeName" />
                            <asp:BoundField DataField="CneeAddr1" HeaderText="Addr 1" ReadOnly="True" SortExpression="CneeAddr1" />
                            <asp:BoundField DataField="CneeTown" HeaderText="Town/City" ReadOnly="True" SortExpression="CneeTown" />
                            <asp:BoundField DataField="CneePostCode" HeaderText="Post Code" ReadOnly="True" SortExpression="CneePostCode" />
                        </Columns>
                        <EmptyDataTemplate>
                            no orders match this filter selection
                        </EmptyDataTemplate>
                        <RowStyle BackColor="#CCFFCC" />
                    </asp:GridView>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
