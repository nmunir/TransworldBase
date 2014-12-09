<%@ Page Language="VB" Theme="AIMSDefault" ValidateRequest="false" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Microsoft.Win32" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Const LOG_CODE_INVENTORYADJUSTMENTSUCCESS As String = "INVENTORYADJUSTMENTSUCCESS"
    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Session("UserKey")) Then
            Server.Transfer("session_expired.aspx")
        End If
        If Not IsPostBack Then
            Call InitCustomerKey()
            Call InitRecentAdjustments()
            Call InitProducts()
            ddlProducts.Focus()
        End If
    End Sub
    
    Protected Sub Log(ByVal sCode As String, ByVal sDescription As String)
        Dim sSQL As String = "INSERT INTO ClientData_CSN_AuditTrail (CreatedOn, Code, Description, CreatedBy) VALUES (GETDATE(), '" & sCode & "', '" & sDescription & "', " & Session("UserKey") & ")"
        Call ExecuteQueryToDataTable(sSQL)
    End Sub
    
    Protected Sub InitProducts()
        Dim sSQL As String = "SELECT ProductCode + ' - ' + ProductDate + ' (' + ProductDescription + ')' 'ProductCode', LogisticProductKey FROM LogisticProduct lp INNER JOIN ClientData_CSN_ProductList pl ON lp.ProductCode = pl.PartNo WHERE lp.CustomerKey = " & pnCustomerKey & " AND DeletedFlag = 'N' ORDER BY lp.ProductCode, lp.ProductDate"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "ProductCode", "LogisticProductKey")
        ddlProducts.Items.Clear()
        ddlProducts.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlProducts.Items.Add(li)
        Next
    End Sub

    Protected Sub InitCustomerKey()
        pnCustomerKey = GetRegistryValue(RegistryHive.LocalMachine, "SOFTWARE\CourierSoftware\CSN", "CSNCustomerKey")
    End Sub
    
    Protected Sub btnAddAdjustment_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call AddAdjustment()
    End Sub
    
    Protected Sub AddAdjustment()
        If ddlReason.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select a reason for the adjustment.")
            ddlReason.Focus()
            Exit Sub
        Else
            If Not IsNumeric(tbAdjustmentQty.Text) Then
                WebMsgBox.Show("Please enter a valid adjustment quantity (eg -10).")
                tbAdjustmentQty.Focus()
                Exit Sub
            Else
                Dim drProduct As DataRow = ExecuteQueryToDataTable("SELECT ProductCode, LanguageId, Misc2 FROM LogisticProduct WHERE LogisticProductKey = " & ddlProducts.SelectedValue).Rows(0)
                If tbReasonOther.Text <> String.Empty Then
                    tbReasonOther.Text = "(" & tbReasonOther.Text.Trim & ")"
                End If
                Dim sbSQL As New StringBuilder
                sbSQL.Append("INSERT INTO ClientData_CSN_InventoryAdjustments")
                sbSQL.Append(" ")
                sbSQL.Append("(")
                sbSQL.Append("CreatedOn")
                sbSQL.Append(",")
                sbSQL.Append("SupplierID")
                sbSQL.Append(",")
                sbSQL.Append("PartNumber")
                sbSQL.Append(",")
                sbSQL.Append("AdjustmentQuantity")
                sbSQL.Append(",")
                sbSQL.Append("SPONumber")
                sbSQL.Append(",")
                sbSQL.Append("Comment")
                sbSQL.Append(",")
                sbSQL.Append("CreatedBy")
                sbSQL.Append(")")
                sbSQL.Append(" ")
                sbSQL.Append("VALUES")
                sbSQL.Append(" ")
                sbSQL.Append("(")
                sbSQL.Append("GETDATE()")
                sbSQL.Append(",")
                sbSQL.Append("'")
                sbSQL.Append(drProduct("Misc2").ToString.Replace("'", "''"))
                sbSQL.Append("'")
                sbSQL.Append(",")
                sbSQL.Append("'")
                sbSQL.Append(drProduct("ProductCode").ToString.Replace("'", "''"))
                sbSQL.Append("'")
                sbSQL.Append(",")
                sbSQL.Append(tbAdjustmentQty.Text)
                sbSQL.Append(",")
                sbSQL.Append("'")
                sbSQL.Append(drProduct("LanguageId").ToString.Replace("'", "''"))
                sbSQL.Append("'")
                sbSQL.Append(",")
                sbSQL.Append("'")
                sbSQL.Append(ddlReason.SelectedValue & ". " & ddlReason.SelectedItem.Text.Replace("'", "''") & " " & tbReasonOther.Text.Replace("'", "''").Replace(Environment.NewLine, " "))
                sbSQL.Append("'")
                sbSQL.Append(",")
                sbSQL.Append(Session("UserKey"))
                sbSQL.Append(")")
                Call ExecuteQueryToDataTable(sbSQL.ToString)
                Call InitRecentAdjustments()

                Call InitProducts()
                ddlReason.SelectedIndex = 0
                tbAdjustmentQty.Text = String.Empty
                tbReasonOther.Text = String.Empty
                'trReasonOther.Visible = False
                btnAddAdjustment.Enabled = False
                ddlProducts.Focus()
                Call Log(LOG_CODE_INVENTORYADJUSTMENTSUCCESS, "Added inventory adjustment " & drProduct("Misc2") & ", " & drProduct("ProductCode") & ", " & drProduct("LanguageId") & ", " & ddlReason.SelectedValue & ")")
            End If
        End If
    End Sub

    Protected Sub InitRecentAdjustments()
        Dim sSQL As String = "SELECT CONVERT(VARCHAR(9), ia.CreatedOn, 6) + ' ' + SUBSTRING((CONVERT(VARCHAR(8), ia.CreatedOn, 108)),1,5) 'AdjustedAt', ia.*, up.FirstName + ' ' + up.LastName + ' (' + up.UserID + ')' 'AdjustedBy' FROM ClientData_CSN_InventoryAdjustments ia INNER JOIN UserProfile up ON ia.CreatedBy = up.[Key] ORDER BY [id]"
        Dim dtRecentAdjustments As DataTable = ExecuteQueryToDataTable(sSQL)
        gvAdjustments.DataSource = dtRecentAdjustments
        gvAdjustments.DataBind()
    End Sub
    
    Protected Sub ddlProducts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            Call InitProduct(ddl.SelectedValue)
            btnAddAdjustment.Enabled = True
        Else
            btnAddAdjustment.Enabled = False
        End If
    End Sub

    Protected Sub InitProduct(ByVal sLogisticProductKey As String)
        Dim sPartNo As String = ExecuteQueryToDataTable("SELECT ProductCode FROM LogisticProduct WHERE LogisticProductKey = " & sLogisticProductKey).Rows(0).Item(0)
        Dim dtProductInfo As DataTable = ExecuteQueryToDataTable("SELECT * FROM ClientData_CSN_ProductList WHERE PartNo = '" & sPartNo & "'")
        If dtProductInfo.Rows.Count = 1 Then
            
        Else
            WebMsgBox.Show("ERROR: Could not locate product info record for product: " & sPartNo & " / " & sLogisticProductKey)
        End If
    End Sub

    Protected Sub lnkbtnRemoveItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Call ExecuteQueryToDataTable("DELETE FROM ClientData_CSN_InventoryAdjustments WHERE [id] = " & lnkbtn.CommandArgument)
        Call InitRecentAdjustments()
        ddlProducts.Focus()
    End Sub

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

    Protected Function GetRegistryValue(ByVal Hive As RegistryHive, ByVal Key As String, ByVal ValueName As String) As String

        Dim objParent As RegistryKey = Nothing
        Dim objSubkey As RegistryKey = Nothing
        Dim sAns As String
        Dim ErrInfo As String = String.Empty

        Select Case Hive
            Case RegistryHive.ClassesRoot
                objParent = Registry.ClassesRoot
            Case RegistryHive.CurrentConfig
                objParent = Registry.CurrentConfig
            Case RegistryHive.CurrentUser
                objParent = Registry.CurrentUser
            Case RegistryHive.LocalMachine
                objParent = Registry.LocalMachine
            Case RegistryHive.PerformanceData
                objParent = Registry.PerformanceData
            Case RegistryHive.Users
                objParent = Registry.Users
        End Select

        Try
            objSubkey = objParent.OpenSubKey(Key)
            'if can't be found, object is not initialized
            If Not objSubkey Is Nothing Then
                sAns = (objSubkey.GetValue(ValueName))
            End If
        Catch ex As Exception
            sAns = "Error"
            'ErrInfo = ex.Message
        Finally
            If ErrInfo = "" And sAns = "" Then
                sAns = "No value found for requested registry key"
            End If
        End Try
        GetRegistryValue = sAns
        
    End Function
    
    Property pnCustomerKey() As Int32
        Get
            Dim o As Object = ViewState("CSNIA_CustomerKey")
            If o Is Nothing Then
                Return -1
            End If
            Return CInt(o)
        End Get
        Set(ByVal Value As Int32)
            ViewState("CSNIA_CustomerKey") = Value
        End Set
    End Property
  
    Protected Sub ddlReason_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Exit Sub
        
        Dim ddl As DropDownList = sender
        If ddl.SelectedValue = "7" Then
            trReasonOther.Visible = True
            tbReasonOther.Text = String.Empty
            tbReasonOther.Focus()
        Else
            trReasonOther.Visible = False
            tbReasonOther.Text = String.Empty
            If tbAdjustmentQty.Text = String.Empty Then
                tbAdjustmentQty.Focus()
            Else
                btnAddAdjustment.Focus()
            End If
        End If
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Wayfair Stock Adjustments</title>
    <style type="text/css">
        .style1
        {
            width: 100%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
      <main:Header id="ctlHeader" runat="server"/>
        <table width="100%" cellpadding="0" cellspacing="0">
            <tr class="bar_accounthandler">
                <td style="width:50%; white-space:nowrap">
                </td>
                <td style="width:50%; white-space:nowrap" align="right">
                </td>
            </tr>
        </table>
            <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="Small" Text="Wayfair Stock Adjustments" Font-Bold="True"/>
            <br />
            <br />
                        <asp:Label ID="Label18" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="INSTRUCTIONS: 1. Select a product 2. Enter the adjustment quantity, plus or minus 3. Select the Reason code 4. Add any useful additional information 5. Click the Add Adjustment button" Font-Bold="True" Font-Italic="True" />
                    <br />
            <br />
        <asp:Panel ID="pnlAdjustment" runat="server" Width="100%" Font-Names="Verdana" Font-Size="XX-Small" GroupingText="Add Adjustment Record">
            <table class="style1">
                <tr>
                    <td style="width: 2%">
                        &nbsp;
                    </td>
                    <td style="width: 20%" align="right">
                        &nbsp;
                        <asp:Label ID="Label15" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="1. Product:" Font-Bold="True" />
                    </td>
                    <td style="width: 2%">
                        &nbsp;
                    </td>
                    <td style="width: 20%">
                        &nbsp;
                        <asp:DropDownList ID="ddlProducts" runat="server" Font-Names="Verdana" Font-Size="XX-Small" onselectedindexchanged="ddlProducts_SelectedIndexChanged" AutoPostBack="True" Width="100%">
                        </asp:DropDownList>
                    </td>
                    <td style="width: 2%">
                        &nbsp;
                    </td>
                    <td style="width: 20%" align="right">
                        &nbsp;
                        <asp:Label ID="Label13" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="2. Adjustment Qty +/-:" />
                    </td>
                    <td style="width: 2%">
                        &nbsp;
                    </td>
                    <td style="width: 20%">
                        &nbsp;<asp:TextBox ID="tbAdjustmentQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Width="50%" MaxLength="5"></asp:TextBox>
                    </td>
                    <td style="width: 2%">
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td>
                    </td>
                    <td align="right">
                        <asp:Label ID="Label11" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="3. Reason:" />
                    </td>
                    <td>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlReason" runat="server" Font-Names="Verdana" Font-Size="XX-Small" style="margin-bottom: 0px" Width="100%" onselectedindexchanged="ddlReason_SelectedIndexChanged" AutoPostBack="True">
                            <asp:ListItem Selected="True" Value="0">- please select -</asp:ListItem>
                            <asp:ListItem Value="1">Adjustment OUT - Damage</asp:ListItem>
                            <asp:ListItem Value="2">Adjustment OUT - Missing</asp:ListItem>
                            <asp:ListItem Value="4">Adjustment OUT - Other</asp:ListItem>
                            <asp:ListItem Value="5">Adjustment IN - Found item</asp:ListItem>
                            <asp:ListItem Value="6">Adjustment IN - Return</asp:ListItem>
                            <asp:ListItem Value="7">Adjustment IN - Other</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td>
                    </td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                    </td>
                    <td>
                        <asp:Button ID="btnAddAdjustment" runat="server" onclick="btnAddAdjustment_Click" Text="Add Adjustment" Enabled="False" />
                    </td>
                    <td>
                    </td>
                </tr>
                <tr id="trReasonOther" runat="server" visible="true">
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        <asp:Label ID="Label19" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="XX-Small" Text="3a. Additional Information (optional):" />
                    </td>
                    <td>
                        &nbsp;</td>
                    <td>
                        <asp:TextBox ID="tbReasonOther" runat="server" Font-Names="Verdana" Font-Size="XX-Small" MaxLength="180" TextMode="MultiLine" Width="100%" Rows="4"></asp:TextBox>
                    </td>
                    <td>
                        &nbsp;</td>
                    <td align="right">
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
            <br />
        </asp:Panel>
        <br />
            <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Recent adjustments:"/>
            <asp:GridView ID="gvAdjustments" runat="server" Width="100%" EnableModelValidation="True" CellPadding="2" Font-Names="Verdana" Font-Size="XX-Small" AutoGenerateColumns="False">
                <Columns>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:LinkButton ID="lnkbtnRemoveItem" runat="server" CommandArgument='<%# Container.DataItem("id")%>' Font-Names="Verdana" Font-Size="XX-Small" OnClick="lnkbtnRemoveItem_Click" OnClientClick="return confirm('Are you sure you want to delete this entry?');">remove</asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="AdjustedAt" HeaderText="Adjusted At" ReadOnly="True" SortExpression="AdjustedAt" />
                    <asp:BoundField DataField="PartNumber" HeaderText="Product Code" ReadOnly="True" SortExpression="PartNumber" />
                    <asp:BoundField DataField="AdjustmentQuantity" HeaderText="Adjustment Qty" ReadOnly="True" SortExpression="AdjustmentQuantity" />
                    <asp:BoundField DataField="Comment" HeaderText="Comment" ReadOnly="True" SortExpression="Comment" />
                    <asp:BoundField DataField="SPONumber" HeaderText="SPO Number" ReadOnly="True" SortExpression="SPONumber" />
                    <asp:BoundField DataField="SupplierID" HeaderText="Supplier ID" ReadOnly="True" SortExpression="SupplierID" />
                    <asp:BoundField DataField="AdjustedBy" HeaderText="Adjusted By" ReadOnly="True" SortExpression="AdjustedBy" />
                </Columns>
                <EmptyDataTemplate>
                    (no adjustments found)
                </EmptyDataTemplate>
        </asp:GridView>
    <p>
                        <asp:Label ID="Label16" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Stock adjustments last sent to Wayfair:" />
                    &nbsp;<asp:Label ID="lblLastUpdateSent" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="True" />
                    </p>
    <p>
                        <asp:Label ID="Label17" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Text="Stock adjustments are sent to Wayfair overnight and cleared from this list." />
                    </p>
    </form>
</body>
</html>
