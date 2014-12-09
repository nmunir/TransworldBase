<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.SqlTypes" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Private gsConn As String = ConfigLib.GetConfigItem_ConnectionString

    Protected Sub Page_Load(ByVal Source As Object, ByVal E As EventArgs)
        If Not IsPostBack Then
            Call PopulateCustomerDropdown()
            Call PopulateWarehouseDropdown()
        End If

    End Sub

    Protected Sub PopulateCustomerDropdown()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' AND ISNULL(AccountHandlerKey, 0) > 0 ORDER BY CustomerAccountCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Text = "DEMO" Then
                ddlCustomer.SelectedIndex = i
                Call PopulateProductDropdown()
                Exit For
            End If
        Next
    End Sub
   
    Protected Sub PopulateWarehouseDropdown()
        Dim sSQL As String = "SELECT WarehouseId, WarehouseKey FROM Warehouse WHERE DeletedFlag = 'N' ORDER BY WarehouseId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseId", "WarehouseKey")
        ddlWarehouse.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlWarehouse.Items.Add(li)
        Next
        For i As Int32 = 0 To ddlWarehouse.Items.Count - 1
            If ddlWarehouse.Items(i).Text = "DEMO" Then
                ddlWarehouse.SelectedIndex = i
                Call InitRackDropdown()
                Exit For
            End If
        Next
        ddlRack.SelectedIndex = 0
        ddlSection.SelectedIndex = 0
        ddlBay.SelectedIndex = 0
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


    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateProductDropdown()
    End Sub

    Protected Sub PopulateProductDropdown()
        Dim sSQL As String = "SELECT LogisticProductKey, ProductCode, ISNULL(ProductDate,'') 'ProductDate', ProductDescription FROM LogisticProduct WHERE DeletedFlag = 'N' AND CustomerKey = " & ddlCustomer.SelectedValue & " " & tbProductQualifier.Text & " ORDER BY ProductCode"
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        ddlProduct.Items.Clear()
        ddlProduct.Items.Add(New ListItem("- please select -", 0))
        For Each dr As DataRow In oDT.Rows
            Dim s As String = dr("ProductCode")
            If dr("ProductDate") <> String.Empty Then
                s += " - " & dr("ProductDate") & " "
            End If
            s += " - " & dr("ProductDescription")
            ddlProduct.Items.Add(New ListItem(s, dr("LogisticProductKey")))
        Next
        lblLogisticProductKey.Text = String.Empty
        gvLocation.Visible = False
    End Sub
    
    Protected Sub ddlWarehouse_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitRackDropdown()
        'Call ClearRackDropdown()
        Call ClearSectionDropdown()
        Call ClearBayDropdown()
        'ddlSection.SelectedIndex = 0
        'ddlBay.SelectedIndex = 0
        ddlRack.Focus()
    End Sub

    Protected Sub InitRackDropdown()
        ddlRack.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseRackId, WarehouseRackKey FROM WarehouseRack WHERE DeletedFlag = 'N' AND WarehouseKey = " & ddlWarehouse.SelectedValue & " ORDER BY WarehouseRackId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseRackId", "WarehouseRackKey")
        ddlRack.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlRack.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlRack_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitSectionDropdown()
        ' Call ClearSectionDropdown()
        Call ClearBayDropdown()
        ddlSection.Focus()
    End Sub

    Protected Sub InitSectionDropdown()
        ddlSection.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseSectionId, WarehouseSectionKey FROM WarehouseSection WHERE DeletedFlag = 'N' AND WarehouseRackKey = " & ddlRack.SelectedValue & " ORDER BY WarehouseSectionId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseSectionId", "WarehouseSectionKey")
        ddlSection.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlSection.Items.Add(li)
        Next
    End Sub
    
    Protected Sub ddlSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call InitBayDropdown()
        ' Call ClearBayDropdown()     ' ??????????????
        ddlSection.Focus()
    End Sub
    
    Protected Sub InitBayDropdown()
        ddlBay.Items.Clear()
        Dim sSQL As String = "SELECT WarehouseBayId, WarehouseBayKey FROM WarehouseBay WHERE DeletedFlag = 'N' AND WarehouseSectionKey = " & ddlSection.SelectedValue & " ORDER BY WarehouseBayId"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "WarehouseBayId", "WarehouseBayKey")
        ddlBay.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlBay.Items.Add(li)
        Next
    End Sub

    Protected Sub ClearRackDropdown()
        If ddlRack.Items.Count > 0 Then
            ddlRack.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearSectionDropdown()
        If ddlSection.Items.Count > 0 Then
            ddlSection.SelectedIndex = 0
        End If
    End Sub
    
    Protected Sub ClearBayDropdown()
        If ddlBay.Items.Count > 0 Then
            ddlBay.SelectedIndex = 0
        End If
        lblBayKey.Text = String.Empty
    End Sub
    
    Protected Sub btnAddQuantity_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsNumeric(tbQtyToAdd.Text) Then
            WebMsgBox.Show("Enter a numeric quantity to add.")
            Exit Sub
        End If
        If CInt(tbQtyToAdd.Text) > 0 Then
            Call AddQuantity()
        Else
            WebMsgBox.Show("Enter a quantity greater than zero.")
        End If
    End Sub
    
    Protected Sub AddQuantity()
        Dim sSQL As String
        sSQL = "SELECT LogisticProductQuantity FROM LogisticProductLocation WHERE LogisticProductKey = " & ddlProduct.SelectedValue & " AND WarehouseBayKey = " & ddlBay.SelectedValue
        Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
        If oDT.Rows.Count > 0 Then
            If oDT.Rows.Count = 1 Then
                Dim nQuantity As Int32 = CInt(oDT.Rows(0).Item(0))
                nQuantity = nQuantity + CInt(tbQtyToAdd.Text)
                If nQuantity >= 0 Then
                    sSQL = "UPDATE LogisticProductLocation SET LogisticProductQuantity = " & nQuantity & " WHERE WarehouseBayKey = " & ddlBay.SelectedValue & " AND LogisticProductKey = " & ddlProduct.SelectedValue
                Else
                    WebMsgBox.Show("Quantity adjustment entered would make total quantity in this location negative.")
                End If
            Else
                WebMsgBox.Show("Error - multiple instances of one product in a single location.")
            End If
        Else
            sSQL = "INSERT INTO LogisticProductLocation (LogisticProductKey, WarehouseBayKey, LogisticProductQuantity, DateStored) VALUES ("
            sSQL += ddlProduct.SelectedValue
            sSQL += ", "
            sSQL += ddlBay.SelectedValue
            sSQL += ", "
            sSQL += tbQtyToAdd.Text
            sSQL += ", GETDATE())"
        End If
        Call ExecuteQueryToDataTable(sSQL)
        tbLog.Text += "Added to " & ddlCustomer.SelectedItem.Text & " " & tbQtyToAdd.Text & " of " & ddlProduct.SelectedItem.Text & " (" & lblLogisticProductKey.Text & ") to Warehouse " & ddlWarehouse.SelectedItem.Text & ", Rack " & ddlRack.SelectedItem.Text & ", Section " & ddlSection.SelectedItem.Text & ", Bay " & ddlBay.SelectedItem.Text & Environment.NewLine
        tbQtyToAdd.Text = String.Empty
        Call SetLocationGrid()
    End Sub
       
    Protected Sub ddlProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call SetLocationGrid()
        If ddlProduct.SelectedIndex > 0 Then
            lblLogisticProductKey.Visible = True
            lblLogisticProductKey.Text = ddlProduct.SelectedValue
        Else
            lblLogisticProductKey.Visible = False
        End If
    End Sub
    
    Protected Sub SetLocationGrid()
        If ddlProduct.SelectedIndex > 0 Then
            gvLocation.Visible = True
            Dim sSQL As String = "SELECT WarehouseId 'Warehouse', WarehouseRackId 'Rack', WarehouseSectionId 'Section', WarehouseBayId 'Bay', LogisticProductQuantity 'Qty' FROM LogisticProductLocation lpl INNER JOIN WarehouseBay wb ON lpl.WarehouseBayKey = wb.WarehouseBayKey INNER JOIN WarehouseSection ws ON wb.WarehouseSectionKey = ws.WarehouseSectionKey INNER JOIN WarehouseRack wr ON ws.WarehouseRackKey = wr.WarehouseRackKey INNER JOIN Warehouse w ON wr.WarehouseKey = w.WarehouseKey WHERE LogisticProductKey = " & ddlProduct.SelectedValue
            Dim oDT As DataTable = ExecuteQueryToDataTable(sSQL)
            gvLocation.DataSource = oDT
            gvLocation.DataBind()
        Else
            gvLocation.Visible = False
        End If
    End Sub
    
    Protected Sub lnkbtnUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hidWarehouse As HiddenField
        Dim hidRack As HiddenField
        Dim hidSection As HiddenField
        Dim hidBay As HiddenField
        Dim gvr As GridViewRow = sender.parent.parent
        hidWarehouse = gvr.FindControl("hidWarehouse")
        hidRack = gvr.FindControl("hidRack")
        hidSection = gvr.FindControl("hidSection")
        hidBay = gvr.FindControl("hidBay")
        For i As Int32 = 0 To ddlWarehouse.Items.Count - 1
            If ddlWarehouse.Items(i).Text = hidWarehouse.Value Then
                ddlWarehouse.SelectedIndex = i
                Exit For
            End If
        Next
        If ddlWarehouse.SelectedIndex > 0 Then
            Call InitRackDropdown()
            For i As Int32 = 0 To ddlRack.Items.Count - 1
                If ddlRack.Items(i).Text = hidRack.Value Then
                    ddlRack.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        If ddlRack.SelectedIndex > 0 Then
            Call InitSectionDropdown()
            For i As Int32 = 0 To ddlSection.Items.Count - 1
                If ddlSection.Items(i).Text = hidSection.Value Then
                    ddlSection.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        If ddlSection.SelectedIndex > 0 Then
            Call InitBayDropdown()
            For i As Int32 = 0 To ddlBay.Items.Count - 1
                If ddlBay.Items(i).Text = hidBay.Value Then
                    ddlBay.SelectedIndex = i
                    Exit For
                End If
            Next
        End If
        tbQtyToAdd.Focus()
    End Sub
    
    Protected Sub lnkbtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Call PopulateProductDropdown()
    End Sub
    
    Protected Sub ddlBay_SelectedIndexChanged(sender As Object, e As System.EventArgs)
        If ddlBay.SelectedIndex = 0 Then
            lblBayKey.Text = String.Empty
        Else
            lblBayKey.Text = "(Bay key: " & ddlBay.SelectedValue & ")"
        End If
    End Sub
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Add Quantity To Product</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
   
    </div>
    <asp:Label ID="Label8" runat="server" Text="WARNING: Only use this utility on the DEMO account, and only add quantity to the DEMO warehouse" Font-Bold="True" ForeColor="Red" Font-Names="Verdana" Font-Size="XX-Small"/>
    <br />
    <br />
    <asp:Label ID="Label1" runat="server" Text="Customer:" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:DropDownList ID="ddlCustomer" runat="server" AutoPostBack="True" onselectedindexchanged="ddlCustomer_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small">
    </asp:DropDownList>
    <br />
    <br />
    <asp:Label ID="Label2" runat="server" Text="Product:" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:DropDownList ID="ddlProduct" runat="server" onselectedindexchanged="ddlProduct_SelectedIndexChanged" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:Label ID="lblLogisticProductKey" runat="server" Visible="False" Font-Names="Verdana" Font-Size="XX-Small"/>
    <br />
    <asp:Label ID="Label7" runat="server" Text="Product Qualifier SQL Clause (eg AND Misc1 &lt;&gt; ''):" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:TextBox ID="tbProductQualifier" runat="server" Width="325px" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:LinkButton ID="lnkbtnSearch" runat="server" onclick="lnkbtnSearch_Click">search</asp:LinkButton>
    <br />
    <br />
    <asp:GridView ID="gvLocation" runat="server" CellPadding="2" Width="100%" Visible="False" EnableModelValidation="True" Font-Names="Verdana" Font-Size="XX-Small">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate><asp:HiddenField ID="hidWarehouse" runat="server"  Value='<%# DataBinder.Eval(Container, "DataItem.Warehouse") %>' /><asp:HiddenField ID="hidRack" runat="server"  Value='<%# DataBinder.Eval(Container, "DataItem.Rack") %>' /><asp:HiddenField ID="hidSection" runat="server"  Value='<%# DataBinder.Eval(Container, "DataItem.Section") %>' /><asp:HiddenField ID="HidBay" runat="server"  Value='<%# DataBinder.Eval(Container, "DataItem.Bay") %>' />
                    <asp:LinkButton ID="lnkbtnUpdate" runat="server" onclick="lnkbtnUpdate_Click">update</asp:LinkButton>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
        <EmptyDataTemplate>
            no product locations
        </EmptyDataTemplate>
    </asp:GridView>
    <br />
    <br />
    <asp:Label ID="Label3" runat="server" Text="Warehouse:" Font-Names="Verdana" Font-Size="XX-Small"/>
    <asp:DropDownList ID="ddlWarehouse" runat="server" AutoPostBack="True" onselectedindexchanged="ddlWarehouse_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:Label ID="Label4" runat="server" Text="Rack:" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:DropDownList ID="ddlRack" runat="server" AutoPostBack="True"
        onselectedindexchanged="ddlRack_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:Label ID="Label5" runat="server" Text="Section:" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:DropDownList ID="ddlSection" runat="server" AutoPostBack="True" onselectedindexchanged="ddlSection_SelectedIndexChanged" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:Label ID="Label6" runat="server" Text="Bay:" Font-Names="Verdana" Font-Size="XX-Small"/>
    &nbsp;<asp:DropDownList ID="ddlBay" runat="server" Font-Names="Verdana" 
        Font-Size="XX-Small" onselectedindexchanged="ddlBay_SelectedIndexChanged" 
        AutoPostBack="True"/>
    &nbsp;
    <asp:Label ID="lblBayKey" runat="server" Font-Names="Verdana" 
        Font-Size="XX-Small"/>
    <br />
    <br />
    <asp:Label ID="Label9" runat="server" Text="Qty to add:" Font-Names="Verdana" Font-Size="XX-Small"/>
    <asp:TextBox ID="tbQtyToAdd" runat="server" Width="87px" Font-Names="Verdana" Font-Size="XX-Small"/>
&nbsp; <asp:Button ID="btnAddQuantity" runat="server" onclick="btnAddQuantity_Click"
        Text="add quantity" />
    <br />
    <br />
    <asp:TextBox ID="tbLog" runat="server" Rows="10" TextMode="MultiLine" Width="100%" Font-Names="Verdana" Font-Size="XX-Small"/>
    </form>
</body>
</html>