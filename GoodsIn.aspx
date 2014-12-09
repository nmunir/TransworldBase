<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient " %>

<%@ Register assembly="Telerik.Web.UI" namespace="Telerik.Web.UI" tagprefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" " http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    ' TO DO
   
    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
   
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("CustomerKey")) Then
            Server.Transfer("session_expired.aspx")
            Exit Sub
        End If
        If Not IsPostBack Then
            Call PopulateCustomerDropdown()
        End If
        Call SetTitle()
    End Sub

    Protected Sub SetTitle()
        Dim sTitle As String = Session("SiteTitle")
        If sTitle <> String.Empty Then
            sTitle += " - "
        End If
        Page.Header.Title = sTitle & "Goods In Journal"
    End Sub
   
    Protected Sub PopulateCustomerDropdown()
        Dim sSQL As String = "SELECT DISTINCT c.CustomerAccountCode, lp.CustomerKey FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE c.DeletedFlag = 'N' AND CustomerStatusId = 'ACTIVE' AND NOT CustomerAccountCode LIKE '%DEMO%' ORDER BY CustomerAccountCode"
        Dim lic As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        rcbCustomer.Items.Clear()
        rcbCustomer.Items.Add(New RadComboBoxItem("- please select - ", 0))
        For Each li As ListItem In lic
            rcbCustomer.Items.Add(New RadComboBoxItem(li.Text, li.Value))
        Next
    End Sub

    Protected Sub PopulateProductDropdown()
        Dim sSQL As String = "SELECT ProductCode + ' ' + ProductDate + ' ' + ProductDescription 'Product', LogisticProductKey FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & rcbCustomer.SelectedValue & " ORDER BY CustomerKey"
        Dim lic As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "Product", "LogisticProductKey")
        rcbProduct.Items.Clear()
        rcbProduct.Items.Add(New RadComboBoxItem("- please select -", 0))
        rcbProduct.Items.Add(New RadComboBoxItem("ALL PRODUCTS", 1))
        For Each li As ListItem In lic
            rcbProduct.Items.Add(New RadComboBoxItem(li.Text, li.Value))
        Next
    End Sub

    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        If b.CommandArgument <> String.Empty Then
            Call InitEdit(b.CommandArgument)
        End If
    End Sub

    Protected Sub InitEdit(nKey As Int32)
        Dim sSQL As String = "SELECT REPLACE(CONVERT(VARCHAR(11), TransactionDateTime, 106), ' ', '-') 'FormattedTransactionDateTime', * FROM GoodsInJournal gij INNER JOIN LogisticProduct lp ON gij.LogisticProductKey = lp.LogisticProductKey WHERE [id] = " & nKey
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        If dt.Rows.Count <> 1 Then
            WebMsgBox.Show("Could not retrieve GoodsInJournal record " & nKey & ". Please inform development.")
        Else
            Dim dr As DataRow = dt.Rows(0)
            rdtpEntryDateTime.Calendar.SelectedDate = dr("FormattedTransactionDateTime")
            rntbQty.Text = dr("GoodsInQty")
            tbComment.Text = dr("Comment")
            hidGoodsInJournalIndex.Value = dr("id")
            lblProductTitle.Text = dr("FormattedTransactionDateTime") & " " & dr("ProductCode") & " " & dr("ProductDescription")
            trProductTitle.Visible = True
            trQty.Visible = True
            trComment.Visible = True
            trButtons.Visible = True
            rntbQty.Focus()
        End If
    End Sub
    
    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim b As Button = sender
        If b.CommandArgument <> String.Empty Then
            Call RemoveRecord(b.CommandArgument)
        End If
    End Sub
   
    Protected Sub RemoveRecord(ByVal nRecord As Integer)
        Call ExecuteQueryToDataTable("DELETE FROM GoodsInJournal WHERE [id] = " & nRecord)
        gvGoodsInJournal.PageIndex = 0
        Call BindGoodsInJournal()
    End Sub
   
    Protected Sub BindGoodsInJournal()
        Dim sSQL As String
        Dim dt As DataTable
        sSQL = "SELECT REPLACE(CONVERT(VARCHAR(11), TransactionDateTime, 106), ' ', '-') 'FormattedTransactionDateTime', * FROM GoodsInJournal gij INNER JOIN LogisticProduct lp ON gij.LogisticProductKey = lp.LogisticProductKey WHERE TransactionDateTime >= GETDATE() - 20 AND gij.CustomerKey = " & rcbCustomer.SelectedValue
        If rcbProduct.SelectedValue <> 1 Then
            sSQL += " AND gij.LogisticProductKey = " & rcbProduct.SelectedValue
        End If
        sSQL += " ORDER BY TransactionDateTime"
        dt = ExecuteQueryToDataTable(sSQL)
        gvGoodsInJournal.DataSource = dt
        gvGoodsInJournal.DataBind()
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
            WebMsgBox.Show("Error in ExecuteQueryToListItemCollection executing: " & sQuery & " : " & ex.Message)
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToListItemCollection = oListItemCollection
    End Function
    
    Protected Sub rcbCustomer_SelectedIndexChanged(sender As Object, e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        If rcbCustomer.SelectedIndex > 0 Then
            Call PopulateProductDropdown()
            trProduct.Visible = True
            If rcbCustomer.Items(0).Value = 0 Then
                rcbCustomer.Items.Remove(0)
            End If
        End If
        gvGoodsInJournal.DataSource = Nothing
        gvGoodsInJournal.DataBind()
        pnlInstructions.Visible = False
    End Sub
    
    Protected Sub rcbProduct_SelectedIndexChanged(sender As Object, e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        If rcbProduct.SelectedValue > 0 Then
            Call BindGoodsInJournal()
            If rcbProduct.SelectedValue <> 1 Then
                lnkbtnAddEntry.Visible = True
            Else
                lnkbtnAddEntry.Visible = False
            End If
            trJournal.Visible = True
            If rcbProduct.Items(0).Value = 0 Then
                rcbProduct.Items.Remove(0)
            End If
        End If
    End Sub
    
    Protected Sub lnkbtnAddEntry_Click(sender As Object, e As System.EventArgs)
        rdtpEntryDateTime.SelectedDate = Date.Now
        trDateTime.Visible = True
        rntbQty.Text = 0
        trQty.Visible = True
        tbComment.Text = String.Empty
        trComment.Visible = True
        trButtons.Visible = True
        rdtpEntryDateTime.Focus()
    End Sub
    
    Protected Sub HideInsertUpdateRows()
        trProductTitle.Visible = False
        trDateTime.Visible = False
        trQty.Visible = False
        trComment.Visible = False
        trButtons.Visible = False
    End Sub
    
    Protected Sub btnSave_Click(sender As Object, e As System.EventArgs)
        Dim sbSQL As New StringBuilder
        If hidGoodsInJournalIndex.Value = String.Empty Then
            sbSQL.Append("INSERT INTO GoodsInJournal (LogisticMovementKey, CustomerKey, LogisticProductKey, TransactionDateTime, GoodsInQty, Comment, LastUpdatedOn, LastUpdatedBy) VALUES (0, ")
            sbSQL.Append(rcbCustomer.SelectedValue)
            sbSQL.Append(", ")
            sbSQL.Append(rcbProduct.SelectedValue)
            sbSQL.Append(", ")
            Dim sDate As String = Format(rdtpEntryDateTime.SelectedDate, "d-MMM-yyyy hh:mm")
            sbSQL.Append("'")
            sbSQL.Append(sDate)
            sbSQL.Append("'")
            sbSQL.Append(", ")
            sbSQL.Append(rntbQty.Text)
            sbSQL.Append(", ")
            sbSQL.Append("'")
            sbSQL.Append(tbComment.Text.Replace("'", "''"))
            sbSQL.Append("'")
            sbSQL.Append(", ")
            sbSQL.Append("GETDATE()")
            sbSQL.Append(", ")
            sbSQL.Append(Session("UserKey"))
            sbSQL.Append(")")
        Else
            sbSQL.Append("UPDATE GoodsInJournal SET GoodsInQty = ")
            sbSQL.Append(rntbQty.Text)
            sbSQL.Append(", ")
            sbSQL.Append("Comment = '")
            sbSQL.Append(tbComment.Text.Replace("'", "''"))
            sbSQL.Append("', LastUpdatedOn = GETDATE(), LastUpdatedBy = ")
            sbSQL.Append(Session("UserKey"))
            sbSQL.Append(" WHERE [id] = ")
            sbSQL.Append(hidGoodsInJournalIndex.Value)
        End If
        Call ExecuteQueryToDataTable(sbSQL.ToString)
        gvGoodsInJournal.PageIndex = 0
        Call BindGoodsInJournal()
        Call HideInsertUpdateRows()
        hidGoodsInJournalIndex.Value = String.Empty
    End Sub

    Protected Sub gvGoodsInJournal_PageIndexChanging(sender As Object, e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gvGoodsInJournal.PageIndex = e.NewPageIndex
        Call BindGoodsInJournal()
    End Sub
    
    Protected Sub btnCancel_Click(sender As Object, e As System.EventArgs)
        hidGoodsInJournalIndex.Value = String.Empty
        Call HideInsertUpdateRows()
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Product Returns</title>
    </head>
<body>
    <form id="form1" runat="server">
     <asp:ScriptManager ID="ScriptManager1" runat="server" />
    <div>
      <main:Header id="ctlHeader" runat="server"/>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                    &nbsp;</td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
        &nbsp;<asp:Label ID="lblTitle" runat="server" Font-Bold="True" Font-Names="Verdana" 
                        Font-Size="Small" Text="Goods In Journal"/>
        &nbsp;<asp:Panel ID="pnlReturn" runat="server" Width="100%" Font-Names="Verdana">
        <table style="width: 100%; font-size: xx-small; font-family: Verdana;" border="0">
            <tr>
                <td align="right" style="width: 10%">
                    Customer:</td>
                <td style="width: 60%">
                    <asp:Panel ID="pnlOrderNumber" runat="server" Width="100%">
                        <telerik:RadComboBox ID="rcbCustomer" Runat="server" AutoPostBack="True" 
                            onselectedindexchanged="rcbCustomer_SelectedIndexChanged">
                        </telerik:RadComboBox>
                    </asp:Panel>
                </td>
                <td style="width: 30%">
                    </td>
            </tr>
            <tr ID="trProduct" runat="server" visible="false">
                <td align="right">
                    Product:</td>
                <td>
                    <telerik:RadComboBox ID="rcbProduct" Runat="server" AutoPostBack="True" 
                        onselectedindexchanged="rcbProduct_SelectedIndexChanged" />
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkbtnAddEntry" runat="server" 
                        onclick="lnkbtnAddEntry_Click" Visible="False">add journal entry</asp:LinkButton>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trUpdate" runat="server" visible="false">
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trProductTitle" runat="server" visible="false">
                <td align="right">
                    &nbsp;</td>
                <td>
                    <asp:Label ID="lblProductTitle" runat="server" Font-Bold="True" 
                        Font-Size="Small" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trDateTime" runat="server" visible="false">
                <td align="right">
                    Date &amp; Time:</td>
                <td>
                    <telerik:RadDateTimePicker ID="rdtpEntryDateTime" Runat="server" 
                        Font-Names="Arial" Font-Size="X-Small">
                    </telerik:RadDateTimePicker>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trQty" runat="server" visible="false">
                <td align="right">
                    Qty:</td>
                <td>
                    <telerik:RadNumericTextBox ID="rntbQty" Runat="server" Font-Names="Arial" 
                        Font-Size="X-Small" MaxValue="1000000" MinValue="1" ShowSpinButtons="True">
                        <NumberFormat DecimalDigits="0" />
                    </telerik:RadNumericTextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trComment" runat="server" visible="false">
                <td align="right">
                    Comment:</td>
                <td>
                    <asp:TextBox ID="tbComment" runat="server" Font-Names="Arial" 
                        Font-Size="X-Small" MaxLength="500" Width="566px"></asp:TextBox>
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trButtons" runat="server" visible="false">
                <td align="right">
                    &nbsp;</td>
                <td>
                    <asp:Button ID="btnSave" runat="server" onclick="btnSave_Click" Text="Save" 
                        Width="100px" />
                    &nbsp;<asp:Button ID="btnCancel" runat="server" onclick="btnCancel_Click" 
                        Text="Cancel" Width="100px" />
                    &nbsp;<asp:HiddenField ID="hidGoodsInJournalIndex" runat="server" />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr ID="trUpdatex" runat="server" visible="false">
                <td align="right">
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr runat="server" id="trJournal" visible="false">
                <td align="right">
                    &nbsp;</td>
                <td colspan="2">
                    <asp:GridView ID="gvGoodsInJournal" runat="server" AllowPaging="True" 
                        AutoGenerateColumns="False" CellPadding="2" EmptyDataText="no entries found" 
                        Font-Names="Verdana" Font-Size="XX-Small" 
                        OnPageIndexChanging="gvGoodsInJournal_PageIndexChanging" Width="100%">
                        <PagerSettings Position="TopAndBottom" />
                        <Columns>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:Button ID="btnEdit" runat="server" 
                                        CommandArgument='<%# DataBinder.Eval(Container.DataItem,"id") %>' 
                                        OnClick="btnEdit_Click" Text="edit" />
                                    <asp:Button ID="btnRemove" runat="server" 
                                        CommandArgument='<%# DataBinder.Eval(Container.DataItem,"id") %>' 
                                        OnClick="btnRemove_Click" 
                                        OnClientClick="return confirm(&quot;Are you sure you want to remove this record?&quot;);" 
                                        Text="remove" />
                                </ItemTemplate>
                                <ItemStyle Wrap="False" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="FormattedTransactionDateTime" HeaderText="Date" 
                                ReadOnly="True" SortExpression="FormattedTransactionDateTime" />
                            <asp:BoundField DataField="ProductCode" HeaderText="Product Code" 
                                ReadOnly="True" SortExpression="ProductCode" />
                            <asp:BoundField DataField="ProductDate" HeaderText="Value / Date" 
                                ReadOnly="True" SortExpression="ProductDate" />
                            <asp:BoundField DataField="ProductDescription" HeaderText="Description" 
                                ReadOnly="True" SortExpression="ProductDescription" />
                            <asp:BoundField DataField="GoodsInQty" HeaderText="Qty" ReadOnly="True" 
                                SortExpression="GoodsInQty" />
                            <asp:BoundField DataField="Comment" HeaderText="Comment" ReadOnly="True" 
                                SortExpression="Comment" />
                        </Columns>
                        <PagerStyle HorizontalAlign="Center" />
                        <EmptyDataTemplate>
                            <asp:Label ID="Label1" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                                Text="no goods in records found" />
                        </EmptyDataTemplate>
                    </asp:GridView>
                </td>
            </tr>
            <tr ID="trRow4" runat="server" visible="false">
                <td align="right">
                    Notes: </td>
                <td>
                    <asp:TextBox ID="tbNotes" runat="server" Font-Names="Verdana" 
                        Font-Size="XX-Small" Width="450px" />
                    &nbsp;(optional)</td>
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

        <asp:Panel runat="server" ID="pnlInstructions" Width="100%">
            <strong>
                <br />
            &nbsp;<asp:Label ID="lblInstructions" runat="server" Font-Bold="True"
                Font-Names="Verdana" Font-Size="Small" Text="Instructions"></asp:Label></strong><table width="100%" >
                <tr>
                    <td style="width:5%">
                        &nbsp;</td>
                    <td style="width:90%">
                        <asp:Label ID="Label12" runat="server" Font-Bold="True" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="#FF5050" style="font-style: italic" 
                            Text="These Goods In entries are copied from the Stock Manager. You can delete or edit an entry (or enter a new entry) to correct the Stock Manager record." />
                    </td>
                    <td style="width:5%">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:Label ID="Label13" runat="server" Font-Bold="True" Font-Names="Verdana" 
                            Font-Size="XX-Small" ForeColor="#FF5050" style="font-style: italic" 
                            Text="WARNING: Events in this journal may be visible to customers!" />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:Label ID="Label16" runat="server" Font-Bold="False" Font-Names="Arial" 
                            Font-Size="XX-Small" ForeColor="Gray" 
                            
                            Text="NOTE: Entries are only available from 1st June 2012 onwards. Only the last 20 days Goods In are shown." />
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        <asp:Label ID="Label14" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="Select the customer for whom Goods In entries are to be modified." />
                        <br />
                        <br />
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="Select a product, or select ALL PRODUCTS." />
                        <br />
                        <br />
                        <asp:Label ID="Label3" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="&lt;b&gt;To add a new Goods In entry&lt;/b&gt;"></asp:Label>
                        <br />
                        <br />
                        <asp:Label ID="Label2" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="1. Select the product for which you want to add an entry.  Click the Add Journal Entry link." />
                        <br />
                        <br />
                        <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="2. Enter values for Date/Time, Quantity and optionally a comment." />
                        <br />
                        <br />
                        <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="3. Click the Save button." />
                        <br />
                        <br />
                        <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="&lt;b&gt;To change an existing Goods In entry&lt;/b&gt;"></asp:Label>
                        <br />
                        <br />
                        <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="1. Click the Edit button next to the entry you want to change." />
                        <br />
                        <br />
                        <asp:Label ID="Label10" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="2. Modify the values." />
                        <br />
                        <br />
                        <asp:Label ID="Label11" runat="server" Font-Names="Verdana" 
                            Font-Size="XX-Small" Text="3. Click the Save button." />
                        <br />
                        <br />
                        <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="&lt;b&gt;To delete a Goods In entry&lt;/b&gt;" />
                        <br />
                        <br />
                        <asp:Label ID="Label7" runat="server" Font-Names="Verdana" Font-Size="XX-Small" 
                            Text="1. Click the Remove button next to the entry you want to delete." />
                        <br />
                        <br />
                        <br />
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