<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient " %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
    ' handle case of removing 2 suppliers (first gets replaced, second doesn't)
    
    Const CUSTOMER_QUANTUM As Int32 = 774
    
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call CleanSupplierList()
            Call CheckForProductsWithNoSupplier()
            Call HideAllPanels()
        End If
        If Not IsUncategorisedSupplier() Then
            btnGenerateReport.Enabled = True
        End If
    End Sub

    Protected Function IsUncategorisedSupplier() As Boolean
        IsUncategorisedSupplier = False
        Dim sSQL As String = "SELECT DISTINCT Misc1 FROM LogisticProduct WHERE CustomerKey = 774 AND DeletedFlag = 'N' AND ArchiveFlag = 'N' AND NOT ProductCode IN ('DVD 1','DVD 2','DVD 3','DVD 4','DVD 5') AND Misc1 NOT IN (SELECT SupplierName FROM ClientData_QL_SupplierType)"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count > 0 Then
            IsUncategorisedSupplier = True
            lblSupplierName.Text = dt.Rows(0).Item(0)
            'cbIsDVDSupplier.Checked = False
            pnlNewSupplier.Visible = True
        Else
            pnlNewSupplier.Visible = False
        End If
    End Function
    
    Protected Sub CleanSupplierList()
        Dim sSQL As String = "DELETE FROM ClientData_QL_SupplierType WHERE SupplierName NOT IN (SELECT DISTINCT Misc1 FROM LogisticProduct WHERE CustomerKey = 774 AND DeletedFlag = 'N')"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub CheckForProductsWithNoSupplier()
        'Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE CustomerKey = 774 AND ISNULL(LTRIM(RTRIM(Misc1)),'') = '' "
        Dim sSQL As String = "SELECT ProductCode FROM LogisticProduct WHERE CustomerKey = 774 AND DeletedFlag = 'N' AND ArchiveFlag = 'N' AND ISNULL(Misc1,'') = '' AND NOT ProductCode IN ('DVD 1','DVD 2','DVD 3','DVD 4','DVD 5')"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        
        If dt.Rows.Count > 0 Then
            Dim sMsg As String = "WARNING! The following products have no supplier code assigned:\n\n"
            For Each dr As DataRow In dt.Rows
                sMsg += dr(0) & "  "
            Next
            WebMsgBox.Show(sMsg)
        End If
    End Sub
    
    Protected Sub HideAllPanels()
        pnlNewSupplier.Visible = False
        pnlSupplierManagement.Visible = False
    End Sub
    
    Protected Sub InitSuppliers()
        Dim sSQL As String = "SELECT [id], SupplierName 'Supplier', IsDVDSupplier FROM ClientData_QL_SupplierType ORDER BY IsDVDSupplier DESC, SupplierName"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvSuppliers.DataSource = dt
        gvSuppliers.DataBind()
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
            'Err.Raise(ex.Message)
            WebMsgBox.Show("Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sIsDVDSupplier As String
        If cbIsDVDSupplier.Checked Then
            sIsDVDSupplier = "1"
        Else
            sIsDVDSupplier = "0"
        End If
        Dim sSQL As String = "INSERT INTO ClientData_QL_SupplierType (SupplierName, IsDVDSupplier) VALUES ('" & lblSupplierName.Text.Replace("'", "''") & "', " & sIsDVDSupplier & ")"
        Call ExecuteQueryToDataTable(sSQL)
        cbIsDVDSupplier.Checked = False
        Call InitSuppliers()
        If Not IsUncategorisedSupplier() Then
            pnlNewSupplier.Visible = False
            btnGenerateReport.Enabled = True
        End If
    End Sub

    Protected Sub lnkbtnToggleSupplierList_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        If lnkbtn.Text.Contains("show") Then
            Call InitSuppliers()
            pnlSupplierManagement.Visible = True
            lnkbtn.Text = "hide supplier list"
        Else
            pnlSupplierManagement.Visible = False
            lnkbtn.Text = "show supplier list"
        End If
    End Sub
    
    Protected Sub lnkbtnToggleHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        If lnkbtn.Text.Contains("show") Then
            pnlHelp.Visible = True
            lnkbtn.Text = "hide help"
        Else
            pnlHelp.Visible = False
            lnkbtn.Text = "show help"
        End If
    End Sub
    
    Protected Sub btnGenerateReport_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsUncategorisedSupplier() Then
            Call GenerateReport()
        Else
            btnGenerateReport.Enabled = False
        End If
    End Sub
    
    Protected Sub GenerateReport()
        Dim sSQL As String = "SELECT DISTINCT ProductCode, ProductDescription, Quantity = CASE ISNUMERIC((SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey)) WHEN 0 THEN 0 ELSE (SELECT SUM(LogisticProductQuantity) FROM LogisticProductLocation AS lpl WHERE lpl.LogisticProductKey = lp.LogisticProductKey) END, Misc1 'Supplier', '_' + LanguageID 'Barcode', 'DVD' 'Category' FROM LogisticProduct AS lp INNER JOIN ClientData_QL_SupplierType qlst ON lp.Misc1 = qlst.SupplierName LEFT OUTER JOIN LogisticProductLocation AS lpl ON lp.LogisticProductKey = lpl.LogisticProductKey WHERE lp.CustomerKey = 774 AND DeletedFlag = 'N' AND IsDVDSupplier = 1 AND NOT ProductCode IN ('DVD 1','DVD 2','DVD 3','DVD 4','DVD 5') ORDER BY Supplier, ProductCode"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        
        Response.Clear()
        Response.ContentType = "text/csv"
        Dim sResponseValue As New StringBuilder
        sResponseValue.Append("attachment; filename=""")
        sResponseValue.Append("QL_DVD_Report_")
        sResponseValue.Append(Format(Date.Now, "ddMMMyyyy_hhmmss"))
        sResponseValue.Append(".csv")
        sResponseValue.Append("""")
        Response.AddHeader("Content-Disposition", sResponseValue.ToString)

        Response.Write("Product Code, Product Description, Quantity In Stock, Supplier, Bar Code, Category" & vbCrLf)
        Dim sItem As String
        For Each dr As DataRow In dt.Rows
            For i = 0 To dt.Columns.Count - 1
                sItem = dr(i) & String.Empty
                If i = 2 Then
                    If CInt(sItem) > 25 Then
                        sItem = "25"
                    End If
                End If
                sItem = sItem.Replace(ControlChars.Quote, ControlChars.Quote & ControlChars.Quote)
                sItem = ControlChars.Quote & sItem & ControlChars.Quote
                Response.Write(sItem)
                Response.Write(",")
            Next
            Response.Write(vbCrLf)
        Next
        Response.End()
    End Sub
    
    Protected Sub lnkbtnRemoveThisSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nID As Int32 = lnkbtn.CommandArgument
        Call ExecuteQueryToDataTable("DELETE FROM ClientData_QL_SupplierType WHERE [id] = " & nID)
        Call InitSuppliers()
        If IsUncategorisedSupplier() Then
            btnGenerateReport.Enabled = False
        End If

    End Sub
    
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Quantum Leap DVD Report</title>
</head>
<body>
    <form id="form1" runat="server">
    <table style="width: 100%">
        <tr>
            <td align="left" style="width: 20%">
                &nbsp;
                <asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="DVD Report"/>
            </td>
            <td style="width: 60%">
                &nbsp;
            </td>
            <td style="width: 20%">
            </td>
        </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnGenerateReport" runat="server" Text="Generate report" Enabled="False" onclick="btnGenerateReport_Click" Width="150px" />
                </td>
                <td>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnToggleSupplierList" runat="server" Font-Size="XX-Small" onclick="lnkbtnToggleSupplierList_Click">show supplier list</asp:LinkButton>
                    &nbsp;
                    <asp:LinkButton ID="lnkbtnToggleHelp" runat="server" Font-Size="XX-Small" onclick="lnkbtnToggleHelp_Click">show help</asp:LinkButton>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
            </tr>
    </table>
    <asp:Panel ID="pnlNewSupplier" runat="server" Width="100%" Font-Names="Verdana">
        <table style="width: 100%">
            <tr>
                <td style="width: 20%">
                </td>
                <td style="width: 60%">
                    <asp:Label ID="lblSupplierName0" runat="server" Font-Size="Small" Font-Italic="True" ForeColor="Maroon">Please categorise the following supplier as a DVD or a non-DVD supplier</asp:Label>
                </td>
                <td style="width: 20%">
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendSupplierName" runat="server" Text="Supplier name:" Font-Size="XX-Small"/>
                    &nbsp;</td>
                <td>
                    <asp:Label ID="lblSupplierName" runat="server" Font-Size="Small"/>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    <asp:Label ID="lblLegendIsDVDSupplier" runat="server" Text="Is a DVD Supplier?:" Font-Size="XX-Small"/>
                </td>
                <td>
                    <asp:CheckBox ID="cbIsDVDSupplier" runat="server" Font-Size="Small" AutoPostBack="True" />
                    &nbsp;
                    <asp:Label ID="lblSupplierName1" runat="server" Font-Italic="True" Font-Size="XX-Small" ForeColor="Maroon">Tick the check box if this supplier is a DVD supplier</asp:Label>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    <asp:Button ID="btnSave" runat="server" Text="Save" Width="80px" onclick="btnSave_Click" />
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlSupplierManagement" runat="server" Width="100%" Font-Names="Verdana">
        <table style="width: 100%">
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
                    &nbsp;
                </td>
                <td>
                    <asp:Label ID="lblSupplierName2" runat="server" Font-Italic="True" Font-Size="XX-Small" ForeColor="Maroon">To change the status of a supplier, remove the supplier; you will be prompted to re-add the supplier, when you can change the status</asp:Label>
&nbsp;<asp:GridView ID="gvSuppliers" runat="server" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" Width="100%" AutoGenerateColumns="False">
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkbtnRemoveThisSupplier" CommandArgument='<%# Container.DataItem("ID")%>' runat="server" onclick="lnkbtnRemoveThisSupplier_Click">remove this supplier</asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="Supplier" HeaderText="Supplier Name" ReadOnly="True" SortExpression="Supplier" />
                            <asp:BoundField DataField="IsDVDSupplier" HeaderText="Is a DVD Supplier?" ReadOnly="True" SortExpression="IsDVDSupplier" />
                        </Columns>
                    </asp:GridView>
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana" Font-Size="Small" Visible="False">
        This report shows Quantum Leap DVD stock levels. If the stock level exceeds 25, it is shown as 25.<br />
        <br />
        The report first checks that all products have a supplier assigned, and warns if one or more products are missing a supplier.<br />
        <br />
        The report next checks that all suppliers are marked as &#39;DVD Supplier&#39; or &#39;not a DVD Supplier&#39;. It prompts for each supplier that is not yet categorised.<br />
        <br />
        Click <em><strong>show supplier list</strong></em> to display the categorised list of suppliers. If a supplier is wrongly categorised, click <em><strong>remove this supplier</strong></em> after which the report will prompt you to add the supplier back in and set the category.<br />
        <br />
        Click the <em><strong>Generate report</strong></em> button to compile the report. If one or more suppliers requires categorisation this button is disabled until you have categorised the missing suppliers.
    </asp:Panel>
    </form>
</body>
</html>
