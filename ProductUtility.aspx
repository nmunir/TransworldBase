<%@ Page Language="VB" Theme="AIMSDefault" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Collections.Generic" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
    
    Dim gsConn As String = ConfigurationSettings.AppSettings("AIMSRootConnectionString")
    Dim gsSQL As String
    Dim gdt As DataTable
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            tbInput.Focus()
        End If
    End Sub

    Protected Function AddConditions() As String
        Dim sClause As String= String.Empty
        If cbUndeletedProductsOnly.Checked Then
            sClause += "lp.DeletedFlag = 'N' AND "
        End If
        If cbUnarchivedProductsOnly.Checked Then
            sClause += "lp.ArchiveFlag = 'N' AND "
        End If
        'sClause.Trimend("AND")
        AddConditions = " " & sClause
    End Function
    
    Protected Sub btnGo_Click(sender As Object, e As System.EventArgs)
        Dim sCommand As String
        sCommand = tbInput.Text.Trim
        If tbInput.Text.StartsWith("#") Then
            sCommand = tbInput.Text.TrimStart("#").Trim
            gsSQL = "SELECT CustomerAccountCode, LogisticProductKey, ProductCode, ProductDescription "
            Call GetFields()
            gsSQL += " FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE " & AddConditions() & " lp.LogisticProductKey = " & sCommand
            gdt = ExecuteQueryToDataTable(gsSQL)
            gvResult.DataSource = gdt
            gvResult.DataBind()
            Exit Sub
        End If
            
        If tbInput.Text.StartsWith("~") Then  ' either ~<CustomerAccountCode>, ~<CustomerKey>, or one of these followed by text to filter on, eg ~HYSTER brochure
            sCommand = tbInput.Text.TrimStart("~").Trim
            Dim nFirstSpace As Int32 = sCommand.IndexOf(" ")
            Dim sCustomer As String = sCommand
            Dim sProductFilter As String = String.Empty
            If nFirstSpace <> -1 Then
                sCustomer = sCommand.Substring(0, nFirstSpace)
                sProductFilter = sCommand.Substring(nFirstSpace).Trim
                sProductFilter = " AND (ProductCode LIKE '%" & sProductFilter & "%' OR ProductDescription LIKE '%" & sProductFilter & "%') "
            End If

            If Not IsNumeric(sCustomer) Then
                Dim nCustomerKey As Int32 = ExecuteQueryToDataTable("SELECT ISNULL(CustomerKey,0) FROM Customer WHERE CustomerAccountCode = '" & sCustomer & "'").Rows(0).Item(0)
                If nCustomerKey > 0 Then
                    gsSQL = "SELECT CustomerAccountCode, LogisticProductKey, ProductCode, ProductDate, ProductDescription "
                    Call GetFields()
                    gsSQL += " FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE " & AddConditions() & " lp.CustomerKey = " & nCustomerKey & sProductFilter & " ORDER BY ProductCode"
                    gdt = ExecuteQueryToDataTable(gsSQL)
                    gvResult.DataSource = gdt
                    gvResult.DataBind()
                Else
                    WebMsgBox.Show("Could not identify customer.")
                End If
            Else
                gsSQL = "SELECT CustomerAccountCode, LogisticProductKey, ProductCode, ProductDate, ProductDescription "
                Call GetFields()
                gsSQL += " FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE " & AddConditions() & " lp.CustomerKey = " & sCustomer & sProductFilter & " ORDER BY ProductCode"
                gdt = ExecuteQueryToDataTable(gsSQL)
                gvResult.DataSource = gdt
                gvResult.DataBind()
            End If
            Exit Sub
        End If
        
        gsSQL = "SELECT CustomerAccountCode, LogisticProductKey, ProductCode, ProductDate, ProductDescription "
        Call GetFields()
        gsSQL += " FROM LogisticProduct lp INNER JOIN Customer c ON lp.CustomerKey = c.CustomerKey WHERE ProductCode LIKE '%" & sCommand & "%' OR ProductDescription LIKE '%" & sCommand & "%' ORDER BY CustomerAccountCode, ProductCode"
        gdt = ExecuteQueryToDataTable(gsSQL)
        gvResult.DataSource = gdt
        gvResult.DataBind()
    End Sub
    
    Protected Sub GetFields()
        If cbProductDepartmentId.Checked Then
            gsSQL += ", lp.ProductDepartmentId"
        End If
        If cbProductDate.Checked Then
            gsSQL += ", lp.ProductDate"
        End If
        If cbLanguageId.Checked Then
            gsSQL += ", lp.LanguageId"
        End If
        If cbItemsPerBox.Checked Then
            gsSQL += ", lp.ItemsPerBox"
        End If
        If cbMinimumStockLevel.Checked Then
            gsSQL += ", lp.MinimumStockLevel"
        End If
        If cbArchiveFlag.Checked Then
            gsSQL += ", lp.ArchiveFlag"
        End If
        If cbSerialNumbersFlag.Checked Then
            gsSQL += ", lp.SerialNumbersFlag"
        End If
        If cbDeletedFlag.Checked Then
            gsSQL += ", lp.DeletedFlag"
        End If
        If cbLastUpdatedByKey.Checked Then
            gsSQL += ", lp.LastUpdatedByKey"
        End If
        If cbLastUpdatedOn.Checked Then
            gsSQL += ", lp.LastUpdatedOn"
        End If
        If cbUnitValue.Checked Then
            gsSQL += ", lp.UnitValue"
        End If
        If cbUnitWeightGrams.Checked Then
            gsSQL += ", lp.UnitWeightGrams"
        End If
        If cbProductCategory.Checked Then
            gsSQL += ", lp.inimumStockLevel"
        End If
        If cbExpiryDate.Checked Then
            gsSQL += ", lp.ExpiryDate"
        End If
        If cbSubCategory.Checked Then
            gsSQL += ", lp.SubCategory"
        End If
        If cbStockOwnedByKey.Checked Then
            gsSQL += ", lp.StockOwnedByKey"
        End If
        If cbMisc1.Checked Then
            gsSQL += ", lp.Misc1"
        End If
        If cbMisc2.Checked Then
            gsSQL += ", lp.Misc2"
        End If
        If cbNotes.Checked Then
            gsSQL += ", lp.Notes"
        End If
        If cbReplenishmentDate.Checked Then
            gsSQL += ", lp.Notes"
        End If
        If cbUnitValueCurrency.Checked Then
            gsSQL += ", lp.ReplenishmentDate"
        End If
        If cbThumbNailImage.Checked Then
            gsSQL += ", lp.ThumbNailImage"
        End If
        If cbWebsiteAdRotatorFlag.Checked Then
            gsSQL += ", lp.WebsiteAdRotatorFlag"
        End If
        If cbOriginalImage.Checked Then
            gsSQL += ", lp.OriginalImage"
        End If
        If cbPDFFileName.Checked Then
            gsSQL += ", lp.PDFFileName"
        End If
        If cbAdRotatorText.Checked Then
            gsSQL += ", lp.AdRotatorText"
        End If
        If cbCreatedOn.Checked Then
            gsSQL += ", lp.CreatedOn"
        End If
        If cbFlag1.Checked Then
            gsSQL += ", lp.Flag1"
        End If
        If cbFlag2.Checked Then
            gsSQL += ", lp.Flag2"
        End If
        If cbInactivityAlertDays.Checked Then
            gsSQL += ", lp.InactivityAlertDays"
        End If
        If cbCalendarManaged.Checked Then
            gsSQL += ", lp.CalendarManaged"
        End If
        If cbOnDemand.Checked Then
            gsSQL += ", lp.OnDemand"
        End If
        If cbOnDemandPriceList.Checked Then
            gsSQL += ", lp.OnDemandPriceList"
        End If
        If cbZeroStockNotification.Checked Then
            gsSQL += ", lp.ZeroStockNotification"
        End If
        If cbCustomLetter.Checked Then
            gsSQL += ", lp.CustomLetter"
        End If
    End Sub
    
    Protected Sub lnkbtnPermissions_Click(sender As Object, e As System.EventArgs)
        Dim lnkbtn As LinkButton = sender
        Dim nLogisticProductKey As Int32 = lnkbtn.CommandArgument
        Dim sSQL As String
        sSQL = "SELECT COUNT (*) FROM UserProductProfile WHERE ProductKey = " & nLogisticProductKey
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        Dim nRecordCount = dt.Rows(0).Item(0)
        If nRecordCount > 1000 Then
        Else
            sSQL = "SELECT up.UserID, up.FirstName, up.LastName, upp.* FROM UserProductProfile upp INNER JOIN UserProfile up ON upp.UserKey = up.[key] WHERE ProductKey = " & nLogisticProductKey & " ORDER BY UserID"
            dt = ExecuteQueryToDataTable(sSQL)
            gvResult2.DataSource = dt
            gvResult2.DataBind()
        End If
    End Sub
    
    Protected Sub btnSelectAll_Click(sender As Object, e As System.EventArgs)
        cbProductDepartmentId.Checked = True
        cbProductDate.Checked = True
        cbLanguageId.Checked = True
        cbItemsPerBox.Checked = True
        cbMinimumStockLevel.Checked = True
        cbArchiveFlag.Checked = True
        cbSerialNumbersFlag.Checked = True
        cbDeletedFlag.Checked = True
        cbLastUpdatedByKey.Checked = True
        cbLastUpdatedOn.Checked = True
        cbUnitValue.Checked = True
        cbUnitWeightGrams.Checked = True
        cbProductCategory.Checked = True
        cbExpiryDate.Checked = True
        cbSubCategory.Checked = True
        cbStockOwnedByKey.Checked = True
        cbMisc1.Checked = True
        cbMisc2.Checked = True
        cbNotes.Checked = True
        cbReplenishmentDate.Checked = True
        cbUnitValueCurrency.Checked = True
        cbThumbNailImage.Checked = True
        cbWebsiteAdRotatorFlag.Checked = True
        cbOriginalImage.Checked = True
        cbPDFFileName.Checked = True
        cbAdRotatorText.Checked = True
        cbCreatedOn.Checked = True
        cbFlag1.Checked = True
        cbFlag2.Checked = True
        cbInactivityAlertDays.Checked = True
        cbCalendarManaged.Checked = True
        cbOnDemand.Checked = True
        cbOnDemandPriceList.Checked = True
        cbZeroStockNotification.Checked = True
        cbCustomLetter.Checked = True
    End Sub

    Protected Sub btnSelectNone_Click(sender As Object, e As System.EventArgs)
        cbProductDepartmentId.Checked = False
        cbProductDate.Checked = False
        cbLanguageId.Checked = False
        cbItemsPerBox.Checked = False
        cbMinimumStockLevel.Checked = False
        cbArchiveFlag.Checked = False
        cbSerialNumbersFlag.Checked = False
        cbDeletedFlag.Checked = False
        cbLastUpdatedByKey.Checked = False
        cbLastUpdatedOn.Checked = False
        cbUnitValue.Checked = False
        cbUnitWeightGrams.Checked = False
        cbProductCategory.Checked = False
        cbExpiryDate.Checked = False
        cbSubCategory.Checked = False
        cbStockOwnedByKey.Checked = False
        cbMisc1.Checked = False
        cbMisc2.Checked = False
        cbNotes.Checked = False
        cbReplenishmentDate.Checked = False
        cbUnitValueCurrency.Checked = False
        cbThumbNailImage.Checked = False
        cbWebsiteAdRotatorFlag.Checked = False
        cbOriginalImage.Checked = False
        cbPDFFileName.Checked = False
        cbAdRotatorText.Checked = False
        cbCreatedOn.Checked = False
        cbFlag1.Checked = False
        cbFlag2.Checked = False
        cbInactivityAlertDays.Checked = False
        cbCalendarManaged.Checked = False
        cbOnDemand.Checked = False
        cbOnDemandPriceList.Checked = False
        cbZeroStockNotification.Checked = False
        cbCustomLetter.Checked = False
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

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body style="font: Verdana; font-size: xx-small">
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="pnlMain" runat="server" Width="100%" DefaultButton="btnGo" Font-Names="Verdana" Font-Size="XX-Small">
            HELP&gt;&gt;&nbsp; text: search product code / description for partial match; #&lt;number&gt;: 
            find product with this key; ~&lt;CustomerKey or Customer Account Code&gt;: list all 
            products for this customer<br />&nbsp;&nbsp;<br />
            <asp:Label ID="Label1" runat="server" Text="Product:" />
            &nbsp;<asp:TextBox ID="tbInput" runat="server" Width="400px" />
            &nbsp;
            <asp:Button ID="btnGo" runat="server" Text="go" onclick="btnGo_Click" />
            &nbsp;<asp:CheckBox ID="cbUndeletedProductsOnly" runat="server" Checked="True" Text="Un-deleted products only" Font-Names="Verdana" Font-Size="XX-Small" />
            &nbsp;<asp:CheckBox ID="cbUnarchivedProductsOnly" runat="server" Text="Un-archived products only" Font-Names="Verdana" Font-Size="XX-Small" />
            <br />
            <br />
            <asp:Button ID="btnSelectAll" runat="server" onclick="btnSelectAll_Click" 
                Text="select all" />
            <asp:Button ID="btnSelectNone" runat="server" Text="select none" 
                onclick="btnSelectNone_Click" />
            &nbsp;<asp:CheckBox ID="cbProductDepartmentId" runat="server" Text="ProductDepartmentId" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbProductDate" runat="server" Text="ProductDate" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbLanguageId" runat="server" Text="LanguageId" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbItemsPerBox" runat="server" Text="ItemsPerBox" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbMinimumStockLevel" runat="server" Text="MinimumStockLevel" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbArchiveFlag" runat="server" Text="ArchiveFlag" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbSerialNumbersFlag" runat="server" Text="SerialNumbersFlag" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbDeletedFlag" runat="server" Text="DeletedFlag" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbLastUpdatedByKey" runat="server" Text="LastUpdatedByKey" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbLastUpdatedOn" runat="server" Text="LastUpdatedOn" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbUnitValue" runat="server" Text="UnitValue" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbUnitWeightGrams" runat="server" Text="UnitWeightGrams" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbProductCategory" runat="server" Text="ProductCategory" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbExpiryDate" runat="server" Text="ExpiryDate" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbSubCategory" runat="server" Text="SubCategory" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbStockOwnedByKey" runat="server" Text="StockOwnedByKey" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbMisc1" runat="server" Text="Misc1" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbMisc2" runat="server" Text="Misc2" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbNotes" runat="server" Text="Notes" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbReplenishmentDate" runat="server" Text="ReplenishmentDate" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbUnitValueCurrency" runat="server" Text="UnitValueCurrency" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbThumbNailImage" runat="server" Text="ThumbNailImage" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbWebsiteAdRotatorFlag" runat="server" Text="WebsiteAdRotatorFlag" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbOriginalImage" runat="server" Text="OriginalImage" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbPDFFileName" runat="server" Text="PDFFileName" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbAdRotatorText" runat="server" Text="AdRotatorText" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbCreatedOn" runat="server" Text="CreatedOn" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbViewOnWebForm" runat="server" Text="ViewOnWebForm" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbRequiresAuthentication" runat="server" Text="RequiresAuthentication" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbRotationProductKey" runat="server" Text="RotationProductKey" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbSubCategory2" runat="server" Text="SubCategory2" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbProductOwner1" runat="server" Text="ProductOwner1" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbProductOwner2" runat="server" Text="ProductOwner2" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbStatus" runat="server" Text="Status" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbUnitValue2" runat="server" Text="UnitValue2" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbFlag1" runat="server" Text="Flag1" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbFlag2" runat="server" Text="Flag2" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbInactivityAlertDays" runat="server" Text="InactivityAlertDays" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbCalendarManaged" runat="server" Text="CalendarManaged" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbOnDemand" runat="server" Text="OnDemand" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbOnDemandPriceList" runat="server" Text="OnDemandPriceList" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbZeroStockNotification" runat="server" Text="ZeroStockNotification" Font-Names="Verdana" Font-Size="XX-Small" />
            <asp:CheckBox ID="cbCustomLetter" runat="server" Text="CustomLetter" Font-Names="Verdana" Font-Size="XX-Small" />
            <br />
        </asp:Panel>
    </div>
    <asp:GridView ID="gvResult" runat="server" Width="100%" CellPadding="2" 
        Font-Names="Verdana" Font-Size="XX-Small">
        <Columns>
            <asp:TemplateField>
                <ItemTemplate>
                    <asp:LinkButton ID="lnkbtnPermissions" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>' onclick="lnkbtnPermissions_Click">permissions</asp:LinkButton>
                    &nbsp;<asp:LinkButton ID="lnkbtnXXX" runat="server" CommandArgument='<%# Container.DataItem("LogisticProductKey")%>'>xxx</asp:LinkButton>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
        <EmptyDataTemplate>
            no data found
        </EmptyDataTemplate>
    </asp:GridView>

    <br />
            <asp:Label ID="Label3" runat="server" Text="Permissions:" />

    <br />

    <asp:GridView ID="gvResult2" runat="server" Width="100%" CellPadding="2" 
        Font-Names="Verdana" Font-Size="XX-Small">
        <EmptyDataTemplate>
            no data found
        </EmptyDataTemplate>
    </asp:GridView>
    </form>
</body>
</html>
