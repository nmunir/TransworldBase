<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
    ' Format return date in Returns History
    ' validate quantity (>0, <= ordered qty)
    ' validate quantity is numeric
    ' put max length on Notes field
    ' check at least one item is to be saved
    ' make history 100% width
    ' select fields to show in product history
    ' provide returns search facility
    ' check return has not already been done for this consignment and warn
   
    'Product Code, Description of Product, Quantity of goods in, Value of Goods, Supplier of Goods in and the date Goods were received.
   
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    'Dim garrMonths() As String = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
   
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            tbOrderNumber.Focus()
            pnlInstructions.Visible = True
            Call PopulateCustomerDropdown()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Returns"
    End Sub
   
    Protected Sub PopulateCustomerDropdown()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerKey IN (SELECT DISTINCT CustomerKey FROM ProductReturns) ORDER BY CustomerAccountCode"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each dr As DataRow In dt.Rows
            ddlCustomer.Items.Add(New ListItem(dr("CustomerAccountCode"), dr("CustomerKey")))
        Next
    End Sub
   
    Protected Sub EnableReturnsHistory()
        If ddlCustomer.Items(0).Value = 0 Then
            ddlCustomer.Items.RemoveAt(0)
        End If
        lnkbtnRefreshReturns.Visible = True
        gvReturnsHistory.Visible = True
    End Sub
   
    Protected Sub AddToCustomerDropdownIfNotPresent(ByVal nCustomerKey As Int32)
        Dim bFound As Boolean = False
        For Each li As ListItem In ddlCustomer.Items
            If li.Value = nCustomerKey Then
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerKey = " & nCustomerKey
            Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
            If dt.Rows.Count <> 1 Then
                WebMsgBox.Show("Error - could not identify Customer Account Code - please notify development.")
            Else
                Dim dr As DataRow = dt.Rows(0)
                ddlCustomer.Items.Add(New ListItem(dr("CustomerAccountCode"), dr("CustomerKey")))
            End If
        End If
        Call EnableReturnsHistory()
    End Sub
   
    Protected Sub SelectCustomerDropdown(ByVal nCustomerKey As Int32)
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Value = nCustomerKey Then
                ddlCustomer.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
   
    Protected Sub HideAllPanels()
        pnlInstructions.Visible = False
        'pnlReturn.Visible = False
        pnlReturns.Visible = False
    End Sub
   
    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        If b.CommandArgument <> String.Empty Then
            Call RemoveRecord(b.CommandArgument)
        End If
    End Sub
   
    Protected Sub RemoveRecord(ByVal nRecord As Integer)
        Call ExecuteQueryToDataTable("DELETE FROM ProductReturns WHERE [id] = " & nRecord)
        gvReturnsHistory.PageIndex = 0
        Call BindReturnsHistory()
    End Sub
   
    Protected Function IsValidReturnData() As Boolean
        IsValidReturnData = True
        If Not IsDate(tbReturnDate.Text) Then
            IsValidReturnData = False
            WebMsgBox.Show("Invalid return date.")
            Exit Function
        End If
        For Each gvr As GridViewRow In gvOrderItems.Rows
            If gvr.RowType = DataControlRowType.DataRow Then
                Dim nOrderQty As Int32 = CInt(gvr.Cells(2).Text)
                Dim tbQtyReturned As TextBox = gvr.FindControl("tbQtyReturned")
                If Not IsNumeric(tbQtyReturned.Text) Then
                    IsValidReturnData = False
                    WebMsgBox.Show("Invalid return quantity.")
                    Exit Function
                End If
                If CInt(tbQtyReturned.Text) > nOrderQty Then
                    IsValidReturnData = False
                    WebMsgBox.Show("Quantity returned cannot exceed original order quantity.")
                    Exit Function
                End If
            End If
        Next
    End Function
   
    Protected Sub BindReturnsGrid()
        Dim sSQL As String = "SELECT pr.*, lp.ProductCode, lp.ProductDescription, REPLACE(CONVERT(VARCHAR(11), ReturnDate, 106), ' ', '-') 'FormattedReturnDate' FROM ProductReturns pr INNER JOIN LogisticProduct lp ON pr.LogisticProductKey = lp.LogisticProductKey WHERE pr.CustomerKey = " & pnCustomerKey & " ORDER BY [id] DESC"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvReturnsHistory.DataSource = dt
        gvReturnsHistory.DataBind()
    End Sub
   
    Protected Sub InitReturnScreen()
        tbOrderNumber.Text = String.Empty
        tbName.Text = String.Empty
        tbReturnDate.Text = String.Empty
        'gvOrderItems.Visible = False
        tbNotes.Text = String.Empty
        trRow1.Visible = False
        trRow2.Visible = False
        trRow3.Visible = False
        trRow4.Visible = False
        trRow5.Visible = False
    End Sub
   
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If IsValidReturnData() Then
            If SaveReturn() Then
                Call AddToCustomerDropdownIfNotPresent(pnCustomerKey)
                Call SelectCustomerDropdown(pnCustomerKey)
                Call EnableReturnsHistory()
                Call InitReturnScreen()
                Call HideAllPanels()
                gvReturnsHistory.PageIndex = 0
                Call BindReturnsGrid()
                pnlReturns.Visible = True
            Else
                WebMsgBox.Show("Could not save entry - please contact your administrator")
            End If
        End If
    End Sub
   
    Protected Function SaveReturn() As Boolean
        Const PREAMBLE As String = "INSERT INTO ProductReturns (CustomerKey, OriginalAWB, LogisticProductKey, ReturnDate, QtyReturned, CneeName, ReturnedToStock, Notes, LastUpdatedBy, LastUpdatedOn) VALUES ("
        For Each gvr As GridViewRow In gvOrderItems.Rows
            Dim hidLogisticProductKey As HiddenField = gvr.FindControl("hidLogisticProductKey")
            Dim tbQtyReturned As TextBox = gvr.FindControl("tbQtyReturned")
            If CInt(tbQtyReturned.Text) > 0 Then
                Dim cbReturnedToStock As CheckBox = gvr.FindControl("cbReturnedToStock")
                Dim nReturnedToStock As Int32 = 0
                If cbReturnedToStock.Checked Then
                    nReturnedToStock = 1
                End If
                Dim sSQL As String = PREAMBLE & pnCustomerKey.ToString & ", '" & tbOrderNumber.Text & "', " & hidLogisticProductKey.Value & ", '" & tbReturnDate.Text & "', " & CInt(tbQtyReturned.Text) & ", '" & tbName.Text.Replace("'", "''") & "', " & nReturnedToStock.ToString & ", '" & tbNotes.Text.Replace("'", "''") & "', " & Session("UserKey") & ", GETDATE())"
                Call ExecuteQueryToDataTable(sSQL)
                SaveReturn = True
            End If
        Next
    End Function
   
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call HideAllPanels()
        pnlReturns.Visible = True
        Call InitReturnScreen()
        pnIsEditingEntry = 0
    End Sub
   
    Protected Function RecordExists(ByVal sOrderNumber As String) As Boolean
        RecordExists = False
        Dim sSQL As String = "SELECT * FROM ProductReturns WHERE OriginalAWB = '" & sOrderNumber.Replace("'", "''") & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        RecordExists = dt.Rows.Count > 0
    End Function
   
    Protected Sub btnFindOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        tbOrderNumber.Text = tbOrderNumber.Text.Trim
        If tbOrderNumber.Text = String.Empty Then
            WebMsgBox.Show("Please enter an order number.")
            tbOrderNumber.Focus()
            Exit Sub
        End If
        Dim sSQL As String = "SELECT * FROM Consignment WHERE AWB = '" & tbOrderNumber.Text.Replace("'", "''") & "'"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count = 0 Then
            WebMsgBox.Show("Could not locate this order - please check the number is correct.")
            tbOrderNumber.Focus()
        ElseIf dt.Rows.Count > 1 Then
            WebMsgBox.Show("Error - more than one order matched. Please inform development.")
            tbOrderNumber.Focus()
        Else
            Dim dr As DataRow = dt.Rows(0)
            pnCustomerKey = dr("CustomerKey")
            Dim sName As String = dr("CneeCtcName") & String.Empty
            If sName = String.Empty Then
                sName = dr("CneeName") & String.Empty
            End If
            tbName.Text = sName
            tbReturnDate.Text = Format(Date.Now, "dd-MMM-yyyy")
            Dim dt2 As DataTable = ExecuteQueryToDataTable("SELECT lm.LogisticProductKey, ProductCode, ProductDescription, ItemsOut FROM LogisticMovement lm INNER JOIN LogisticProduct lp ON lm.LogisticProductKey = lp.LogisticProductKey WHERE ConsignmentKey = " & dr("key"))
            gvOrderItems.DataSource = dt2
            gvOrderItems.DataBind()
            trRow1.Visible = True
            trRow2.Visible = True
            trRow3.Visible = True
            trRow4.Visible = True
            trRow5.Visible = True
            Call SelectCustomerDropdown(pnCustomerKey)
            Call EnableReturnsHistory()
            gvReturnsHistory.PageIndex = 0
            Call BindReturnsGrid()
            If RecordExists(tbOrderNumber.Text) Then
                WebMsgBox.Show("WARNING: One or more returns have already been recorded for this consignment.")
            End If
        End If
    End Sub

    Protected Sub lnkbtnRefreshReturns_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCustomer.SelectedIndex >= 0 Then
            If ddlCustomer.SelectedValue > 0 Then
                Call BindReturnsHistory()
            End If
        End If
    End Sub

    Protected Sub BindReturnsHistory()
        Dim sSQL As String = "SELECT pr.*, lp.ProductCode, lp.ProductDescription, REPLACE(CONVERT(VARCHAR(11), ReturnDate, 106), ' ', '-') 'FormattedReturnDate' FROM ProductReturns pr INNER JOIN LogisticProduct lp ON pr.LogisticProductKey = lp.LogisticProductKey WHERE pr.CustomerKey = " & ddlCustomer.SelectedValue & " ORDER BY [id]"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        gvReturnsHistory.DataSource = dt
        gvReturnsHistory.DataBind()
    End Sub
   
    Protected Sub gvReturns_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvReturnsHistory.PageIndex = e.NewPageIndex
        Call BindReturnsHistory()
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

    Property pbIsEditingEntry() As Boolean
        Get
            Dim o As Object = ViewState("RET_IsEditingEntry")
            If o Is Nothing Then
                Return False
            End If
            Return CBool(o)
        End Get
        Set(ByVal Value As Boolean)
            ViewState("RET_IsEditingEntry") = Value
        End Set
    End Property

    Property pnIsEditingEntry() As Integer
        Get
            Dim o As Object = ViewState("RET_IsEditingEntry")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Integer)
            ViewState("RET_IsEditingEntry") = Value
        End Set
    End Property

    Property pnCustomerKey() As Int32
        Get
            Dim o As Object = ViewState("RET_CustomerKey")
            If o Is Nothing Then
                Return 0
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("RET_CustomerKey") = Value
        End Set
    End Property

    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call EnableReturnsHistory()
        gvReturnsHistory.PageIndex = 0
        Call BindReturnsHistory()
    End Sub
   
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Product Returns</title>
    </head>
<body>
    <form id="form1" runat="server">
    <div>
      <main:Header id="ctlHeader" runat="server"></main:Header>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                    &nbsp;</td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="left" style="width: 20%">
        &nbsp;<asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Returns"></asp:Label>
        </td>
                <td style="width: 60%">
                    &nbsp;</td>
                <td style="width: 20%">
                </td>
            </tr>
        </table>
        <asp:Panel ID="pnlReturn" runat="server" Width="100%" Font-Names="Verdana">
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="right" style="width: 10%">
                    Order #:</td>
                <td style="width: 60%">
                    <asp:Panel ID="pnlOrderNumber" runat="server" Width="100%" DefaultButton="btnFindOrder">
                    <asp:TextBox ID="tbOrderNumber" runat="server" Font-Names="Verdana" Font-Size="Small" Width="120px"></asp:TextBox>
                    &nbsp;<asp:Button ID="btnFindOrder" runat="server" Text="Find Order" onclick="btnFindOrder_Click" />
                    </asp:Panel>
                </td>
                <td style="width: 30%">
                </td>
            </tr>
            <tr runat="server" id="trRow1" visible="false">
                <td align="right">
                    Name:</td>
                <td>
                    <asp:TextBox ID="tbName" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="200px"></asp:TextBox>
                    &nbsp;&nbsp;&nbsp;&nbsp; Return Date:
                    <asp:TextBox ID="tbReturnDate" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="120px"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr runat="server" id="trRow2" visible="false">
                <td align="right">
                    Order:</td>
                <td>
                    <asp:GridView ID="gvOrderItems" runat="server" Width="100%" CellPadding="2" AutoGenerateColumns="False">
                        <Columns>
                            <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductCode" />
                            <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                            <asp:BoundField DataField="ItemsOut" HeaderText="Order Qty" ReadOnly="True" SortExpression="ItemsOut" />
                            <asp:TemplateField HeaderText="Qty Returned">
                                <ItemTemplate>
                                    <asp:TextBox ID="tbQtyReturned" runat="server" Font-Names="Verdana" Font-Size="Small" Width="40px">0</asp:TextBox>
                                    <asp:HiddenField ID="hidLogisticProductKey" Value='<%# DataBinder.Eval(Container, "DataItem.LogisticProductKey") %>' runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Returned to stock?">
                                <ItemTemplate>
                                    <asp:CheckBox ID="cbReturnedToStock" runat="server" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
                <td>
                    <strong>&nbsp;<br />&nbsp;&nbsp;</strong>&nbsp;</td>
            </tr>
            <tr runat="server" id="trRow3" visible="false">
                <td align="right">
                </td>
                <td>
                    &nbsp;</td>
                <td>
                </td>
            </tr>
            <tr runat="server" id="trRow4" visible="false">
                <td align="right">
                    Notes:
                </td>
                <td>
                    <asp:TextBox ID="tbNotes" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="450px"/>
                    &nbsp;(optional)</td>
                    <td>
                    </td>
            </tr>
            <tr runat="server" id="trRow5" visible="false">
                <td align="right">
                </td>
                <td>
                    <asp:Button ID="btnSave" runat="server" Text="save" Width="100px"
                        onclick="btnSave_Click" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" onclick="btnCancel_Click"
                        Text="cancel" Width="100px" CausesValidation="False" />
                </td>
                <td>
                </td>
            </tr>
            <tr>
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
        </table>
        </asp:Panel>

        <asp:Panel ID="pnlReturns" runat="server" Width="100%">
            <strong>
                &nbsp;</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br />
            <table style="width: 100%">
            <tr>
            <td style="width: 30%">
                <strong>
                <asp:Label ID="lblUploadHistory" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="Small" Text="Returns History" />
                </strong>
            </td>
            <td style="width: 30%">
                <asp:LinkButton ID="lnkbtnRefreshReturns" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onclick="lnkbtnRefreshReturns_Click" Text="refresh" Visible="False" />
            </td>
            <td style="width: 30%" align="right">
                Customer:
                <asp:DropDownList ID="ddlCustomer" runat="server" AutoPostBack="True" onselectedindexchanged="ddlCustomer_SelectedIndexChanged">
                </asp:DropDownList>
            </td>
            </tr>
            </table>
            <table style="width: 100%">
                <tr>
                    <td style="width: 100%">
                        <asp:GridView ID="gvReturnsHistory" runat="server" AllowPaging="True" CellPadding="2" EmptyDataText="no entries found" Font-Names="Verdana" Font-Size="XX-Small" OnPageIndexChanging="gvReturns_PageIndexChanging" Width="100%" Visible="False" AutoGenerateColumns="False">
                            <PagerSettings Position="TopAndBottom" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:Button ID="btnRemove" runat="server" CommandArgument='<%# DataBinder.Eval(Container.DataItem,"id") %>' OnClick="btnRemove_Click" Text="remove" OnClientClick='return confirm("Are you sure you want to remove this record?");'/>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="OriginalAWB" HeaderText="AWB" ReadOnly="True" SortExpression="OriginalAWB" />
                                <asp:BoundField DataField="ProductCode" HeaderText="Product Code" ReadOnly="True" SortExpression="ProductCode" />
                                <asp:BoundField DataField="ProductDescription" HeaderText="Description" ReadOnly="True" SortExpression="ProductDescription" />
                                <asp:BoundField DataField="FormattedReturnDate" HeaderText="Return Date" ReadOnly="True" SortExpression="ReturnDate" />
                                <asp:BoundField DataField="QtyReturned" HeaderText="Qty Returned" ReadOnly="True" SortExpression="QtyReturned" />
                                <asp:BoundField DataField="CneeName" HeaderText="Name" ReadOnly="True" SortExpression="CneeName" />
                                <asp:BoundField DataField="ReturnedToStock" HeaderText="Returned to stock?" ReadOnly="True" SortExpression="ReturnedToStock" />
                                <asp:BoundField DataField="Notes" HeaderText="Notes" ReadOnly="True" SortExpression="Notes" />
                            </Columns>
                            <PagerStyle HorizontalAlign="Center" />
                            <EmptyDataTemplate>
                                <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="no returns found" />
                            </EmptyDataTemplate>
                        </asp:GridView>
                        <br />
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel runat="server" ID="pnlInstructions" Visible="false" Width="100%">
            <strong>
                <br />
            &nbsp;<asp:Label ID="lblInstructions" runat="server" Font-Bold="True"
                Font-Names="Verdana" Font-Size="Small" Text="Instructions"></asp:Label></strong><table width="100%" >
                <tr>
                    <td style="width:5%">
                        &nbsp;</td>
                    <td style="width:90%">
                        &nbsp;</td>
                    <td style="width:5%">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="<b>To record a returned product</b>"></asp:Label><br />
                        <br />
                        <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="1. Enter the original consignment number." /><br />
                        <br />
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="2. For each product returned, enter the number of items returned."/><br />
                        <br />
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="3. Check the name and date is correct (the date is used on the product status report)."/><br />
                        <br />
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="4. Add notes if required."/><br />
                        <br />
                        <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="5. Click the <b>save</b> button. NOTE: this utility does NOT return items to stock - you must do that separately."/>
                        .<br /><br />
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="<b>To view the list of returned products</b>"/><br />
                        <br />
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="1. From the dropdown blox select the customer to view (only customers with one or more return records are listed)." /><br />
                        <br />
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="2. To remove a return record click the <b>remove</b> button."/><br />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        </div>
    </form>
</body>
</html>