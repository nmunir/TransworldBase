<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="main" TagName="Header" Src="main_header.ascx" %>
<%@ Register assembly="Telerik.Web.UI" namespace="Telerik.Web.UI" tagprefix="telerik" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    ' TO DO
    ' validate switch customer params
    ' check rcbModelUser has a valid user selected
    
    Const ALL_USERS As String = "9999"
    Const ITEMS_PER_REQUEST As Integer = 30

    Dim gsConn As String = ConfigLib.GetConfigItem_ConnectionString
    'Dim goConn As New SqlConnection(gsConn)
    'Dim oCmd As SqlCommand
    'Dim sbDetail As New StringBuilder, sbSummary As New StringBuilder, sbSQL As New StringBuilder
    'Dim nTotalUsers As Integer = 0, nUsersWithMissingProducts As Integer = 0

    Dim gnTimeout As Int32
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not IsPostBack Then
            Call BindCustomerList()
            Call BindNewCustomerList()
        End If
    End Sub
    
    Protected Sub btnSwitchCustomer_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ValidateSwitchCustomerParameters() Then
            
        End If
    End Sub
    
    Protected Function ValidateSwitchCustomerParameters() As Boolean
        ValidateSwitchCustomerParameters = True
        If Not (IsNumeric(rcbUser.SelectedValue) AndAlso rcbUser.SelectedValue > 0) Then
            WebMsgBox.Show("Please select the user who you want to switch from one account to another..")
            ValidateSwitchCustomerParameters = False
            Exit Function
        End If
        If ddlModelUser.Enabled Then
            If ddlModelUser.SelectedIndex = 0 Then
                WebMsgBox.Show("a message.")
                ValidateSwitchCustomerParameters = False
                Exit Function
            End If
        End If
        If rcbModelUser.Enabled Then
            If rcbModelUser.SelectedIndex = 0 Then
                WebMsgBox.Show("a message.")
                ValidateSwitchCustomerParameters = False
                Exit Function
            End If
        End If
    End Function

    Protected Sub rcbModelUser_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim data As DataTable = GetModelUsers(e.Text)
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbModelUser.DataTextField = "UserID"
        rcbModelUser.DataValueField = "UserKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            rcb.Text = data.Rows(i)("UserID").ToString()
            rcb.Value = data.Rows(i)("UserKey").ToString()
            rcbModelUser.Items.Add(rcb)
        Next
        e.Message = GetStatusMessage(endOffset, data.Rows.Count)
    End Sub

    Protected Function GetModelUsers(Optional ByVal sFilter As String = "") As DataTable
        Dim sFilterClause As String = String.Empty
        If sFilter <> String.Empty Then
            sFilter = sFilter.Replace("'", "''")
            sFilterClause += " UserID LIKE '" & sFilter & "%' AND "
        End If
        Dim sSQL As String = "SELECT UserID + ' (' + FirstName + ' ' + LastName + ')' 'UserID', [key] 'UserKey' FROM UserProfile WHERE Status = 'Active' AND Type = 'User' AND " & sFilterClause & " CustomerKey = " & ddlNewCustomer.SelectedValue & " ORDER BY UserID"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        GetModelUsers = dt
    End Function

    Protected Sub rcbModelUser_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        'If rcb.SelectedIndex <= 0 Then
        '    btnSwitchCustomer.Enabled = False
        'Else
        '    btnSwitchCustomer.Enabled = True
        'End If
        If rcb.Text = String.Empty Then
            btnSwitchCustomer.Enabled = False
        Else
            btnSwitchCustomer.Enabled = True
        End If
    End Sub

    Protected Sub ddlNewCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddlModelUser.Enabled = False
        ddlModelUser.Visible = False
        rcbModelUser.Enabled = False
        rcbModelUser.Visible = False
        lblLegendWUModelUser.Visible = False
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            If ddl.SelectedValue = 579 Then
                ddlModelUser.Enabled = True
                ddlModelUser.Visible = True
                lblLegendWUModelUser.Visible = True
                btnSwitchCustomer.Enabled = False
                Call PopulateModelUserDropdownForWURS()
            ElseIf ddl.SelectedValue = 686 Then
                ddlModelUser.Enabled = True
                ddlModelUser.Visible = True
                lblLegendWUModelUser.Visible = True
                btnSwitchCustomer.Enabled = False
                Call PopulateModelUserDropdownForWUIRE()
            Else
                rcbModelUser.Enabled = True
                rcbModelUser.Visible = True
                rcbModelUser.Text = String.Empty
                btnSwitchCustomer.Enabled = False
                lblLegendWUModelUser.Visible = True
            End If
        End If
        btnSwitchCustomer.Enabled = False
    End Sub

    Protected Sub rcbUser_ItemsRequested(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxItemsRequestedEventArgs)
        Dim data As DataTable = GetUsers(e.Text)
        Dim itemOffset As Integer = e.NumberOfItems
        Dim endOffset As Integer = Math.Min(itemOffset + ITEMS_PER_REQUEST, data.Rows.Count)
        e.EndOfItems = endOffset = data.Rows.Count
        rcbUser.DataTextField = "UserID"
        rcbUser.DataValueField = "UserKey"
        For i As Int32 = itemOffset To endOffset - 1
            Dim rcb As New RadComboBoxItem
            rcb.Text = data.Rows(i)("UserID").ToString()
            rcb.Value = data.Rows(i)("UserKey").ToString()
            rcbUser.Items.Add(rcb)
        Next
        e.Message = GetStatusMessage(endOffset, data.Rows.Count)
    End Sub

    Protected Function GetUsers(Optional ByVal sFilter As String = "") As DataTable
        Dim sFilterClause As String = String.Empty
        If sFilter <> String.Empty Then
            sFilter = sFilter.Replace("'", "''")
            sFilterClause += " UserID LIKE '" & sFilter & "%' AND "
        End If
        Dim sSQL As String = "SELECT UserID + ' (' + FirstName + ' ' + LastName + ')' 'UserID', [key] 'UserKey' FROM UserProfile WHERE Status = 'Active' AND Type = 'User' AND " & sFilterClause & " CustomerKey = " & ddlCustomer.SelectedValue & " ORDER BY UserID"
        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
        GetUsers = dt
    End Function

    Private Shared Function GetStatusMessage(ByVal nOffset As Integer, ByVal nTotal As Integer) As String
        If nTotal <= 0 Then
            Return "No matches"
        End If
        'Return [String].Format("Items <b>1</b>-<b>{0}</b> of <b>{1}</b>", nOffset, nTotal)
        If nOffset <= ITEMS_PER_REQUEST Then
            'GetStatusMessage = "Click for more items " & nOffset & " " & nTotal & " " & ITEMS_PER_REQUEST
            GetStatusMessage = "Click for more items"
        End If
        'GetStatusMessage = "Click for more items" 
        If nOffset = nTotal Then
            GetStatusMessage = "No more items"
        Else
            GetStatusMessage = "Click for more items"
        End If
    End Function
    
    Protected Sub rcbUser_SelectedIndexChanged(ByVal o As Object, ByVal e As Telerik.Web.UI.RadComboBoxSelectedIndexChangedEventArgs)
        Dim rcb As RadComboBox = o
        'Dim ddl As DropDownList = sender
        If rcb.SelectedIndex = 0 Then
            'btnAddToOrder.Enabled = False
            'lnkbtnPlus1.Enabled = False
            'lnkbtnPlus5.Enabled = False
            'lnkbtnMinus1.Enabled = False
            'lnkbtnMinus5.Enabled = False
            'rntbQty.Enabled = False
            'tbQty.Enabled = False
        Else
            'btnAddToOrder.Enabled = True
            'lnkbtnPlus1.Enabled = True
            'lnkbtnPlus5.Enabled = True
            'lnkbtnMinus1.Enabled = True
            'lnkbtnMinus5.Enabled = True
            'tbQty.Enabled = True
            'tbQty.Focus()
            'rntbQty.Enabled = True
            'rntbQty.Focus()
        End If
        'Call ClearOrderConfirmation()
    End Sub

    Protected Sub BindCustomerList()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlCustomer.Items.Clear()
        ddlCustomer.Items.Add(New ListItem("- all customers -", 0))
        For Each li As ListItem In oListItemCollection
            ddlCustomer.Items.Add(li)
        Next
    End Sub

    Protected Sub BindNewCustomerList()
        Dim sSQL As String = "SELECT CustomerAccountCode, CustomerKey FROM Customer WHERE CustomerStatusId = 'ACTIVE' ORDER BY CustomerAccountCode"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "CustomerAccountCode", "CustomerKey")
        ddlNewCustomer.Items.Clear()
        ddlNewCustomer.Items.Add(New ListItem("- select a customer -", 0))
        For Each li As ListItem In oListItemCollection
            ddlNewCustomer.Items.Add(li)
        Next
    End Sub

    Protected Sub BindProductDropdown()
        Dim sSQL As String = "SELECT ProductCode + ' ' + ISNULL(ProductDate,'') + ' - ' + ProductDescription 'Product', LogisticProductKey FROM LogisticProduct WHERE CustomerKey = " & ddlCustomer.SelectedValue & " AND ArchiveFlag = 'N' AND DeletedFlag = 'N' ORDER BY Product"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "Product", "LogisticProductKey")
        ddlProduct.Items.Clear()
        ddlProduct.Items.Add(New ListItem("- please select -", 0))
        For Each li As ListItem In oListItemCollection
            ddlProduct.Items.Add(li)
        Next
    End Sub

    Protected Sub BindUserGroupDropdown()
        Dim sSQL As String = "SELECT GroupName, id FROM UP_UserPermissionGroups WHERE CustomerKey = " & ddlCustomer.SelectedValue & "  ORDER BY GroupName"
        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "GroupName", "id")
        ddlUserGroup.Items.Clear()
        ddlUserGroup.Items.Add(New ListItem("- please select -", 0))
        ddlUserGroup.Items.Add(New ListItem("- all users -", ALL_USERS))
        For Each li As ListItem In oListItemCollection
            ddlUserGroup.Items.Add(li)
        Next
    End Sub

    Protected Sub lnkbtnWURS_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Text = "WURS" Then
                ddlCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        Call BindProductDropdown()
        Call BindUserGroupDropdown()
        rcbUser.Enabled = True
        rcbUser.Text = String.Empty
    End Sub

    Protected Sub lnkbtnWUIRE_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlCustomer.Items.Count - 1
            If ddlCustomer.Items(i).Text = "WUIRE" Then
                ddlCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        Call BindProductDropdown()
        Call BindUserGroupDropdown()
        rcbUser.Enabled = True
        rcbUser.Text = String.Empty
    End Sub

    Protected Sub lnkbtnNewCustomerWURS_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlNewCustomer.Items.Count - 1
            If ddlNewCustomer.Items(i).Text = "WURS" Then
                ddlNewCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        rcbModelUser.Enabled = False
        rcbModelUser.Visible = False
        ddlModelUser.Enabled = True
        ddlModelUser.Visible = True
        lblLegendWUModelUser.Visible = True
        Call PopulateModelUserDropdownForWURS()
    End Sub

    Protected Sub lnkbtnNewCustomerWUIRE_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        For i As Int32 = 0 To ddlNewCustomer.Items.Count - 1
            If ddlNewCustomer.Items(i).Text = "WUIRE" Then
                ddlNewCustomer.SelectedIndex = i
                Exit For
            End If
        Next
        rcbModelUser.Enabled = False
        rcbModelUser.Visible = False
        ddlModelUser.Enabled = True
        ddlModelUser.Visible = True
        lblLegendWUModelUser.Visible = True
        Call PopulateModelUserDropdownForWUIRE()
    End Sub

    Protected Sub ddlProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then

        End If
        Call InitialiseNewPermissions()
        ddlUserGroup.SelectedIndex = 0
    End Sub
    
    Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Call BindProductDropdown()
        Call BindUserGroupDropdown()
        Call InitialiseNewPermissions()
        If ddlCustomer.SelectedIndex > 0 Then
            rcbUser.Enabled = True
        Else
            rcbUser.Enabled = False
            rcbUser.Text = String.Empty
        End If
    End Sub
    
    Protected Sub rbAllowedToPick_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            cbApplyMaxGrab.Enabled = True
            tbMaxGrabQty.Text = String.Empty
            cbApplyMaxGrab.Focus()
            btnUpdateProductPermissions.Enabled = True
        End If
    End Sub

    Protected Sub rbNotAllowedToPick_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rb As RadioButton = sender
        If rb.Checked Then
            cbApplyMaxGrab.Enabled = False
            cbApplyMaxGrab.Checked = False
            tbMaxGrabQty.Text = String.Empty
            btnUpdateProductPermissions.Enabled = True
        End If
    End Sub

    Protected Sub cbApplyMaxGrab_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cb As CheckBox = sender
        If cb.Checked Then
            tbMaxGrabQty.Enabled = True
            tbMaxGrabQty.Focus()
        Else
            tbMaxGrabQty.Enabled = False
        End If
        tbMaxGrabQty.Text = String.Empty
    End Sub
    
    Protected Sub InitialiseNewPermissions()
        rbAllowedToPick.Checked = False
        rbNotAllowedToPick.Checked = False
        cbApplyMaxGrab.Checked = False
        cbApplyMaxGrab.Enabled = False
        tbMaxGrabQty.Text = String.Empty
        tbMaxGrabQty.Enabled = False
        btnUpdateProductPermissions.Enabled = False
    End Sub
    
    Protected Sub btnUpdateProductPermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.Validate()
        If Page.IsValid Then
            If ValidateProductPermissionsUpdate() Then
                Call UpdateProductPermissions()
            End If
        End If
    End Sub

    Protected Sub UpdateProductPermissions()
        Dim sAbleToPick As String = "0"
        Dim sApplyMaxGrab As String = "0"
        Dim sMaxGrabQty As String = "0"
        If rbAllowedToPick.Checked Then
            sAbleToPick = "1"
        End If
        If cbApplyMaxGrab.Checked Then
            sApplyMaxGrab = "1"
        End If
        If IsNumeric(tbMaxGrabQty.Text) Then
            sMaxGrabQty = tbMaxGrabQty.Text
        End If
        Dim sUpdateClause As String = "UPDATE UserProductProfile SET AbleToPick = " & sAbleToPick & ", ApplyMaxGrab = " & sApplyMaxGrab & ", MaxGrabQty = " & sMaxGrabQty & " WHERE ProductKey = " & ddlProduct.SelectedValue & " AND UserKey IN "
        Dim sUserSelectionClause As String = String.Empty
        Dim sConfirmationMessage As String = "Updated product permissions for customer " & ddlCustomer.SelectedItem.Text & ", product " & ddlProduct.SelectedItem.Text
        If ddlUserGroup.SelectedValue = ALL_USERS Then
            sUserSelectionClause = "(SELECT [key] FROM UserProfile WHERE Status = 'Active' AND Type = 'User' AND DeletedFlag = 0 AND CustomerKey = " & ddlCustomer.SelectedValue & ")"
            sConfirmationMessage += ", to ALL USERS"
        Else
            sUserSelectionClause = "(SELECT [key] FROM UserProfile up INNER JOIN UP_UserPermissionGroups upg ON up.UserGroup = upg.[id] WHERE Status = 'Active' AND Type = 'User' AND DeletedFlag = 0 AND upg.[id] = " & ddlUserGroup.SelectedValue & ")"
            sConfirmationMessage += ", to users in User Group " & ddlUserGroup.SelectedItem.Text
        End If
        sConfirmationMessage += ".\n\nNew permissions: "
        If rbAllowedToPick.Checked Then
            sConfirmationMessage += "Can Pick, "
            If cbApplyMaxGrab.Checked Then
                sConfirmationMessage += "MAX GRAB enabled, with a value of " & tbMaxGrabQty.Text & "."
            Else
                sConfirmationMessage += "MAX GRAB disabled (any quantity allowed)."
            End If
        Else
            sConfirmationMessage += "Cannot Pick."
        End If
        Dim sSQL As String = sUpdateClause & sUserSelectionClause & " SELECT @@ROWCOUNT"
        Dim nRowcount As Int32 = ExecuteQueryToDataTable(sSQL).Rows(0).Item(0)
        WebMsgBox.Show(sConfirmationMessage.Replace("  ", " ") & "\n\n" & nRowcount.ToString & " user(s) updated.")
    End Sub
    
    Protected Function ValidateProductPermissionsUpdate() As Boolean
        ValidateProductPermissionsUpdate = True
        tbMaxGrabQty.Text = tbMaxGrabQty.Text.Trim
        If cbApplyMaxGrab.Checked AndAlso tbMaxGrabQty.Text = String.Empty Then
            WebMsgBox.Show("Please enter a value for MAX GRABS.")
            ValidateProductPermissionsUpdate = False
            Exit Function
        End If
        If cbApplyMaxGrab.Checked AndAlso CInt(tbMaxGrabQty.Text) = 0 Then
            WebMsgBox.Show("You have entered a value of 0 for MAX GRABS. Either enter a value greater than 0 or remove the Apply MAX GRAB attribute.")
            ValidateProductPermissionsUpdate = False
            Exit Function
        End If
        If ddlUserGroup.SelectedIndex = 0 Then
            WebMsgBox.Show("Please select a user group, or 'all users'.")
            ValidateProductPermissionsUpdate = False
            Exit Function
        End If
    End Function
    
    Protected Sub PopulateModelUserDropdownForWURS()
        ddlModelUser.Items.Clear()
        ddlModelUser.Items.Add(New ListItem("- please select -", 0))
        ddlModelUser.Items.Add(New ListItem("ALBEMARLE AND BOND - 5KEN - 13508", 13508))
        ddlModelUser.Items.Add(New ListItem("CASH CONVERTERS - A06V - 6313", 6313))
        ddlModelUser.Items.Add(New ListItem("CASH EXPRESS - A1KZ - 6324", 6324))
        ddlModelUser.Items.Add(New ListItem("CASH GENERATOR - 2PW4 - 15791", 15791))
        ddlModelUser.Items.Add(New ListItem("THE CO-OPERATIVE TRAVEL - 9DB7 - 14710", 14710))
        ddlModelUser.Items.Add(New ListItem("EUROCHANGE - JH2G - 9894", 9894))
        ddlModelUser.Items.Add(New ListItem("EXCHANGE INTERNATIONAL - C371 - 6813", 6813))
        ddlModelUser.Items.Add(New ListItem("HARVEY & THOMPSON - A12G - 17063", 17063))
        ddlModelUser.Items.Add(New ListItem("HERBERT BROWN - 5NTJ - 13529", 13529))
        ddlModelUser.Items.Add(New ListItem("ICE PLC (HUMBERSIDE AIRPORT) - M0FL - 7431", 7431))
        ddlModelUser.Items.Add(New ListItem("KANOO - OR9L - 7536", 7536))
        ddlModelUser.Items.Add(New ListItem("NCA - U6GI - 7927", 7927))
        ddlModelUser.Items.Add(New ListItem("RAMSDENS - STWN - 8144", 8144))
        ddlModelUser.Items.Add(New ListItem("THE MONEY SHOP - 1EC3 - 15863", 15863))
        ddlModelUser.Items.Add(New ListItem("THE TRADE EXCHANGE - DGTW - 8747", 8747))
        ddlModelUser.Items.Add(New ListItem("CHEQUE CASHING - DCG9 - 6459", 6459))
        ddlModelUser.Items.Add(New ListItem("THE TRAVEL HOUSE - MGXD - 8750", 8750))
        ddlModelUser.Items.Add(New ListItem("TRAVELCARE - EMN4 - 9336", 9336))
        ddlModelUser.Items.Add(New ListItem("FIRST CHOICE TRAVEL - A774 - 7003", 7003))
        ddlModelUser.Items.Add(New ListItem("CHEQUE CENTRE - JGKH - 11364", 11364))
        ddlModelUser.Items.Add(New ListItem("PREPAID (NEWSMARK PECKHAM) - AWYA - 7972", 7972))
        ddlModelUser.Items.Add(New ListItem("INDEPENDENT (ABI FOOD) - A4NG - 17238", 17238))
    End Sub
    
    Protected Sub PopulateModelUserDropdownForWUIRE()
        ddlModelUser.Items.Clear()
        ddlModelUser.Items.Add(New ListItem("- please select -", 0))
        ddlModelUser.Items.Add(New ListItem("WUIRE LOW AGENTS (MALLOW CREDIT UNION) - 7TKA - 12555", 12555))
        ddlModelUser.Items.Add(New ListItem("WUIRE HIGH AGENTS (IRELANDS EDUCATION) - 2OQ7 - 15282", 15282))
        ddlModelUser.Items.Add(New ListItem("WUIRE MEDIUM AGENTS (CASH CREATORS) -  - 15810", 15810))
        ddlModelUser.Items.Add(New ListItem("WUIRE HIGH TRANSACTOR (CALL @ NET STONEY BATTER) - LV2Q - 12778", 12778))
        ddlModelUser.Items.Add(New ListItem("WUIRE TOP AGENT (ANTECH MUNSTER) - G20C - 12671", 12671))
        ddlModelUser.Items.Add(New ListItem("WUIRE STAFF - aidan.kennerk - 12833", 12833))
    End Sub
    
    Protected Sub ddlModelUser_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddl As DropDownList = sender
        If ddl.SelectedIndex > 0 Then
            btnSwitchCustomer.Enabled = True
        Else
            btnSwitchCustomer.Enabled = False
        End If
    End Sub
    
    'Protected Sub btnVerify_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call Verify()
    'End Sub
        
    'Protected Function ConvertUserIDs() As String
    '    Dim sUserIDs As String = tbUserIDs.Text.Trim
    '    sUserIDs = sUserIDs.Replace(Environment.NewLine, " ")
    '    sUserIDs = sUserIDs.Replace(",", " ")
    '    ConvertUserIDs = sUserIDs
    'End Function
    
    'Protected Sub Verify(Optional ByVal bVerifyUserInDropdown As Boolean = False)
    '    If Not IsNumeric(tbQuitLevel.Text) Then
    '        lblError.Text = "Value must be numeric"
    '        Exit Sub
    '    End If
        
    '    Dim sSQL As String

    '    If cbOutputDebugInformation.Checked Then
    '        Call ExecuteQueryToDataTable("DELETE FROM AAA_Debug")
    '    End If

    '    sSQL = "SELECT * FROM UserProfile WHERE Type = 'User'"
    '    If ddlCustomer.SelectedIndex > 0 Then
    '        sSQL += " AND CustomerKey = " & ddlCustomer.SelectedValue
    '    End If
    '    If Not cbIncludeSuspendedUsers.Checked Then
    '        sSQL += " AND Status = 'Active'"
    '    End If
    '    If ddlUser.SelectedIndex > 0 Then
    '        sSQL += " AND [key] = " & ddlUser.SelectedValue
    '    End If
        
    '    Dim sUserIDs() As String
    '    tbUserIDs.Text = tbUserIDs.Text.Trim
    '    If ConvertUserIDs() <> String.Empty Then
    '        sUserIDs = ConvertUserIDs.Split(" ")
    '        Dim sUserList As String = String.Empty
    '        For Each s As String In sUserIDs
    '            If s.Length > 0 Then
    '                Dim sType As String = String.Empty
    '                Try
    '                    sType = ExecuteQueryToDataTable("SELECT Type FROM UserProfile WHERE UserID = '" & s & "'").Rows(0).Item(0) & String.Empty
    '                Catch
    '                End Try
    '                If sType = String.Empty Then
    '                    WebMsgBox.Show("Could not find User ID " & s)
    '                    Exit Sub
    '                Else
    '                    If sType.ToLower <> "user" Then
    '                        WebMsgBox.Show("Account " & s & " is a " & sType & " account, not a standard User account. Please remove from list.")
    '                        Exit Sub
    '                    End If
    '                    sUserList += "'" & s & "', "
    '                End If
    '            End If
    '        Next
    '        sUserList += "''"
    '        sSQL = "SELECT * FROM UserProfile WHERE Type = 'User' AND UserID IN (" & sUserList & ")"
    '    End If
    '    If bVerifyUserInDropdown AndAlso ddlUser.SelectedIndex > 0 Then
    '        sSQL = "SELECT * FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue
    '    End If

    '    If cbOutputDebugInformation.Checked Then
    '        Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('SQL: " & sSQL.Replace("'", "''") & "')")
    '    End If

    '    Dim oAdapter1 As New SqlDataAdapter(sSQL, goConn)
    '    Dim tblUsers As New DataTable
    '    Dim nRecordCount As Integer = 0
    '    oAdapter1.Fill(tblUsers)
        
    '    For Each drUser As DataRow In tblUsers.Rows
    '        Dim tblMissingProducts As DataTable
    '        nTotalUsers += 1
            
    '        Dim nUserKey As Integer = drUser.Item("Key")
    '        Dim nCustomerKey As Integer = drUser.Item("CustomerKey")
    '        Dim oTable As DataTable = GetProductCountForCustomer(nCustomerKey)
    '        Dim nExpectedProductCount As Integer = oTable.Rows(0).Item(0)
    '        If cbOutputDebugInformation.Checked Then
    '            Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('" & nTotalUsers.ToString & " UserID: " & drUser("UserID") & "')")
    '        End If
    '        Dim nMissingProductCount As Integer
    '        tblMissingProducts = GetMissingProductsForUser(nUserKey, nCustomerKey)
    '        nMissingProductCount = tblMissingProducts.Rows.Count

    '        If nMissingProductCount > 0 And Not (cbDontReportWhenEntireProfileMissing.Checked And nMissingProductCount = nExpectedProductCount) Then
    '            If cbOutputDebugInformation.Checked Then
    '                Call ExecuteQueryToDataTable("INSERT INTO AAA_Debug (Result) VALUES ('Products missing')")
    '            End If
    '            nUsersWithMissingProducts += 1
    '            If cbGenerateSQLToFixMissingEntries.Checked Then
    '                sbSQL.Append("PRINT 'Restoring values for user " & drUser.Item("UserId") & " (" & drUser.Item("Key") & ")'" & vbNewLine & vbNewLine)
    '            End If
    '            For Each drMissingProduct As DataRow In tblMissingProducts.Rows
    '                sbDetail.Append("UserID: " & drUser.Item("UserId") & "(" & drUser.Item("Key") & ") missing product " & drMissingProduct.Item("ProductCode") & "(" & drMissingProduct.Item("LogisticProductKey") & ")" & vbNewLine)
                    
    '                If cbGenerateSQLToFixMissingEntries.Checked Then
    '                    sbSQL.Append("INSERT INTO UserProductProfile (")
    '                    sbSQL.Append("UserKey, ProductKey, AbleToView, AbleToPick, AbleToEdit, AbleToArchive, AbleToDelete, ApplyMaxGrab, MaxGrabQty) ")
    '                    sbSQL.Append(" VALUES (")
    '                    sbSQL.Append(drUser.Item("Key") & ", " & drMissingProduct.Item("LogisticProductKey") & ", 1, 1, 1, 1, 1, 0, 0)" & vbNewLine)
    '                End If
    '            Next
    '            If cbGenerateSQLToMakeActive.Checked Then
    '                sbSQL.Append("UPDATE UserProfile SET Status = 'Active' WHERE Type = 'User' AND UserID = '" & drUser.Item("UserId") & "'" & vbNewLine)
    '            End If
    '            sbSummary.Append(tblMissingProducts.Rows.Count.ToString & " item(s) out of " & nExpectedProductCount & " missing for " & GetCustomerName(nCustomerKey.ToString) & " user " & drUser.Item("UserId") & "(" & drUser.Item("Key") & ")" & vbNewLine)
    '            If cbGenerateSQLToFixMissingEntries.Checked Then
    '                sbSQL.Append(vbNewLine)
    '            End If
    '        End If
            
    '        If nUsersWithMissingProducts >= CInt(tbQuitLevel.Text) Then
    '            Exit For
    '        End If
            
    '    Next
    '    sbSummary.Append("Total users checked so far: " & nTotalUsers.ToString & vbNewLine)
    '    sbSummary.Append("Total users found with missing products: " & nUsersWithMissingProducts.ToString & vbNewLine)
    '    tbSummary.Text = sbSummary.ToString
    '    tbDetail.Text = sbDetail.ToString
    '    tbSQL.Text = sbSQL.ToString
        
    '    If bVerifyUserInDropdown AndAlso ddlUser.SelectedIndex > 0 Then
    '        sSQL = "SELECT * FROM UserProductProfile WHERE UserKey = " & ddlUser.SelectedValue
    '        Dim dt As DataTable = ExecuteQueryToDataTable(sSQL)
    '        gvProductProfile.Visible = True
    '        lnkbtnHideProductProfile.Visible = True
    '        gvProductProfile.DataSource = dt
    '        gvProductProfile.DataBind()
    '    End If

    'End Sub
    
    'Protected Function GetCustomerName(ByVal sCustomerKey As String) As String
    '    Dim sSQL As String
    '    sSQL = "SELECT CustomerAccountCode FROM Customer WHERE CustomerKey = " & sCustomerKey
    '    Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
    '    Dim tblCustomerName As New DataTable
    '    oAdapter.Fill(tblCustomerName)
    '    GetCustomerName = tblCustomerName.Rows(0).Item(0)
    '    oAdapter.Dispose()
    '    tblCustomerName.Dispose()
    'End Function
    
    'Protected Function GetMissingProductsForUser(ByVal nUserKey As Integer, ByVal nCustomerKey As Integer) As DataTable
    '    Dim sSQL As String
    '    sSQL = "SELECT * FROM LogisticProduct WHERE CustomerKey = " & nCustomerKey.ToString & " "
    '    sSQL += "AND NOT LogisticProduct.LogisticProductKey IN "
    '    sSQL += "(SELECT ProductKey FROM UserProductProfile WHERE UserKey = " & nUserKey.ToString & ")"
    '    Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
    '    Dim tblMissingProducts As New DataTable
    '    oAdapter.Fill(tblMissingProducts)
    '    Return tblMissingProducts
    'End Function
    
    'Protected Function GetProductCountForCustomer(ByVal nCustomerKey As Integer) As DataTable
    '    Dim sSQL As String = "SELECT COUNT(*) FROM LogisticProduct WHERE CustomerKey = " & nCustomerKey.ToString
    '    Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
    '    Dim tblProductCount As New DataTable
    '    oAdapter.Fill(tblProductCount)
    '    Return tblProductCount
    'End Function
    
    'Protected Function GetProductProfileForUser(ByVal nUserKey As Integer) As DataTable
    '    Dim sSQL As String = "SELECT * FROM UserProductProfile WHERE UserKey = " & nUserKey.ToString
    '    Dim oAdapter As New SqlDataAdapter(sSQL, goConn)
    '    Dim tblUserProductProfile As New DataTable
    '    oAdapter.Fill(tblUserProductProfile)
    '    Return tblUserProductProfile
    'End Function
    
    'Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    tbSummary.Text = ""
    '    tbDetail.Text = ""
    '    tbSQL.Text = ""
    '    gvProductProfile.Visible = False
    'End Sub
    
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
            'tbSQL.Text += "Error in ExecuteQueryToDataTable executing: " & sQuery & " : " & ex.Message
        Finally
            oConn.Close()
        End Try
        ExecuteQueryToDataTable = oDataTable
    End Function
    
    'Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call PopulateUserDropdown()
    'End Sub
    
    'Protected Sub PopulateUserDropdown()
    '    ddlModelUser.SelectedIndex = 0
    '    ddlUser.Items.Clear()
    '    lnkbtnVerifySelectedUser.Enabled = False
    '    Dim sSQL As String
    '    If ddlCustomer.SelectedIndex > 0 Then
    '        sSQL = "SELECT UserID + ' (' + FirstName + ' ' + LastName + ') - ' + ISNULL(GroupName,'') UserName, [key]  FROM UserProfile up LEFT OUTER JOIN UP_UserPermissionGroups upg on up.UserGroup = upg.[id] WHERE up.CustomerKey = " & ddlCustomer.SelectedValue
    '        If Not cbIncludeSuspendedUsers.Checked Then
    '            sSQL += " AND Status = 'Active'"
    '        End If
    '        sSQL += " ORDER BY UserID"
    '        Dim oListItemCollection As ListItemCollection = ExecuteQueryToListItemCollection(sSQL, "UserName", "Key")
    '        ddlUser.Items.Clear()
    '        ddlUser.Items.Add(New ListItem("- all users -", 0))
    '        For Each li As ListItem In oListItemCollection
    '            ddlUser.Items.Add(li)
    '        Next
    '        lnkbtnVerifySelectedUser.Enabled = True
    '    End If
    'End Sub
    
    'Protected Sub btnCloneProfiles_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call CloneProfiles()
    'End Sub
    
    'Protected Function GetUserIDFromUserKey(ByVal nUserKey As Int32) As String
    '    GetUserIDFromUserKey = String.Empty
    '    Dim sSQL As String = "SELECT UserID FROM UserProfile WHERE [key] = " & nUserKey
    '    Dim dtUserProfile As DataTable = ExecuteQueryToDataTable(sSQL)
    '    If dtUserProfile.Rows.Count = 1 Then
    '        GetUserIDFromUserKey = dtUserProfile.Rows(0).Item(0)
    '    Else
    '        WebMsgBox.Show("GetUserIDFromUserKey: could not find record!")
    '    End If
    'End Function
    
    'Protected Sub CloneProfiles()
    '    Dim sSQL As String = String.Empty
    '    Dim sUserIDs() As String

    '    If ddlUser.SelectedIndex = -1 Then
    '        WebMsgBox.Show("Please select source customer.")
    '        Exit Sub
    '    End If
    '    If Not ddlUser.SelectedValue > 0 Then
    '        WebMsgBox.Show("Please select source account.")
    '        Exit Sub
    '    End If
    '    Dim sSourceUserID As String = GetUserIDFromUserKey(ddlUser.SelectedValue)
    '    ' sSQL = "SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue
    '    Dim nSourceUserGroup = ExecuteQueryToDataTable("SELECT ISNULL(UserGroup, 0) FROM UserProfile WHERE [key] = " & ddlUser.SelectedValue).Rows(0).Item(0)
    '    tbUserIDs.Text = tbUserIDs.Text.Trim
    '    If ConvertUserIDs.Trim <> String.Empty Then
    '        sUserIDs = ConvertUserIDs.Split(" ")
    '        Dim sUserList As String = String.Empty
    '        For Each s As String In sUserIDs
    '            If s.Length > 0 Then
    '                Dim sType As String = String.Empty
    '                Dim dtUser As DataTable = ExecuteQueryToDataTable("SELECT Type, CustomerKey, ISNULL(UserGroup, 0) 'UserGroup' FROM UserProfile WHERE UserID = '" & s & "'")
    '                If dtUser.Rows.Count <> 1 Then
    '                    WebMsgBox.Show("Could not find User ID " & s)
    '                    Exit Sub
    '                ElseIf s = sSourceUserID Then
    '                    WebMsgBox.Show("Destination UserID list contains source UserID! Please remove.")
    '                    Exit Sub
    '                ElseIf Not cbOverrideCloningGroupSanityCheck.Checked Then
    '                    If nSourceUserGroup <> dtUser.Rows(0).Item("UserGroup") Then
    '                        WebMsgBox.Show("Account " & s & " has a different user group (" & dtUser.Rows(0).Item("UserGroup") & ") to that of source (" & nSourceUserGroup.ToString & ")")
    '                        Exit Sub
    '                    End If
    '                End If
    '                sType = dtUser.Rows(0).Item("Type")
    '                If dtUser.Rows(0).Item("Type").ToString.ToLower <> "user" Then
    '                    WebMsgBox.Show("Account " & s & " is a " & sType & " account, not a standard User account. Please remove from list.")
    '                    Exit Sub
    '                End If
    '                If dtUser.Rows(0).Item("CustomerKey") <> ddlCustomer.SelectedValue Then
    '                    WebMsgBox.Show("Source account and destination account are for different customers.")
    '                    Exit Sub
    '                End If
    '                sUserList += "'" & s & "', "
    '            End If
    '        Next
    '        sUserList += "''"
            
    '        sSQL = "SELECT * FROM UserProfile WHERE Type = 'User' AND UserID IN (" & sUserList & ")"
    '        Dim dtDestinationUsers As DataTable = ExecuteQueryToDataTable(sSQL)
    '        For Each drDestinationUser As DataRow In dtDestinationUsers.Rows
    '            Dim nDestinationUserKey As Integer = drDestinationUser.Item("Key")
    '            Call CloneSingleProfile(ddlUser.SelectedValue, nDestinationUserKey)
    '            tbDetail.Text += "Cloned " & drDestinationUser("UserID") & " from " & ddlUser.SelectedItem.Text & vbCrLf
    '        Next
    '    End If
    'End Sub

    'Protected Sub CloneSingleProfile(ByVal nSourceUser As Int32, ByVal nDestinationUser As Int32)
    '    Dim sbSQL As New StringBuilder
    '    sbSQL.Append("DELETE FROM UserProductProfile WHERE UserKey = ")
    '    sbSQL.Append(nDestinationUser.ToString)
    '    sbSQL.Append(" ")
    '    sbSQL.Append("INSERT INTO UserProductProfile (UserKey,ProductKey,AbleToView,AbleToPick,AbleToEdit,AbleToArchive,AbleToDelete,ApplyMaxGrab,MaxGrabQty)")
    '    sbSQL.Append(" ")
    '    sbSQL.Append("SELECT ")
    '    sbSQL.Append(nDestinationUser.ToString)
    '    sbSQL.Append(", ProductKey, AbleToView, AbleToPick, AbleToEdit, AbleToArchive, AbleToDelete, ApplyMaxGrab, MaxGrabQty FROM UserProductProfile WHERE UserKey = ")
    '    sbSQL.Append(nSourceUser.ToString)
    '    Call ExecuteQueryToDataTable(sbSQL.ToString)
    'End Sub

    'Protected Sub lnkbtnBuildSQLString_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call BuildSQLString()
    'End Sub
    
    'Protected Sub BuildSQLString()
    '    Dim sbElementList As New StringBuilder
    '    Dim sUserIDs() As String

    '    tbUserIDs.Text = tbUserIDs.Text.Trim
    '    sUserIDs = ConvertUserIDs.Split(" ")
    '    sbElementList.Append("(")
    '    For Each s As String In sUserIDs
    '        sbElementList.Append("'" & s & "', ")
    '    Next
    '    sbElementList.Append(")")
    '    tbDetail.Text = sbElementList.ToString.Replace(", )", ")")
    'End Sub
    
    'Protected Sub lnkbtnToggleHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    pnlHelp.Visible = Not pnlHelp.Visible
    'End Sub
    
    'Protected Sub lnkbtnVerifySelectedUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call Verify(True)
    'End Sub
    
    'Protected Sub lnkbtnClearUserIDListbox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    tbUserIDs.Text = String.Empty
    'End Sub
    
    'Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs)
    '    gnTimeout = Server.ScriptTimeout
    '    Server.ScriptTimeout = 3600
    'End Sub

    'Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Server.ScriptTimeout = gnTimeout
    'End Sub
    
    'Protected Sub cbOverrideCloningGroupSanityCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Dim cb As CheckBox = sender
    '    If cb.Checked Then
    '        cb.ForeColor = Drawing.Color.Red
    '        cb.Font.Bold = True
    '    Else
    '        cb.ForeColor = Drawing.Color.Empty
    '        cb.Font.Bold = false
    '    End If
    'End Sub
    
    'Protected Sub lnkbtnHideProductProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    lnkbtnHideProductProfile.Visible = False
    '    gvProductProfile.Visible = False
    'End Sub

</script>
<html xmlns=" http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Product Configurator</title>
    <style type="text/css">
        .style1
        {
            height: 24px;
        }
        .style2
        {
            height: 28px;
        }
    </style>
</head>
<body style="font-family: Verdana">
    <form id="frmProductConfigurator" runat="server">
    <telerik:RadScriptManager ID="rsm" runat="server" />
    <main:Header ID="ctlHeader" runat="server"/>
    <table style="width: 100%" cellpadding="0" cellspacing="0">
        <tr class="bar_noticeboard">
            <td style="width: 50%; white-space: nowrap" />
            <td style="width: 50%; white-space: nowrap" align="right" />
        </tr>
    </table>
    <strong>Product Configurator - build 30JUL12A</strong><br />
    <table style="width: 100%">
        <tr>
            <td style="width: 1%" />
            <td style="width: 11%" />
            <td style="width: 32%" />
            <td style="width: 23%" />
            <td style="width: 32%" />
            <td style="width: 1%" />
        </tr>
        <tr>
            <td>
            </td>
            <td>
                <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="Customer" />
            </td>
            <td>
                <asp:DropDownList ID="ddlCustomer" runat="server" Font-Size="X-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged">
                    <asp:ListItem Value="-1">- select a customer -</asp:ListItem>
                </asp:DropDownList>
                &nbsp;<asp:LinkButton ID="lnkbtnWURS" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnWURS_Click">WURS</asp:LinkButton>
                &nbsp;<asp:LinkButton ID="lnkbtnWUIRE" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnWUIRE_Click">WUIRE</asp:LinkButton>
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td colspan="5">
                <hr />
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="2">
                <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="Modify Product Permissions" />
            </td>
            <td>
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
            <td>
            </td>
            <td>
                &nbsp;
            </td>
            <td colspan="2">
                <asp:DropDownList ID="ddlProduct" runat="server" Font-Size="X-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlProduct_SelectedIndexChanged">
                    <asp:ListItem Value="-1">- select a product -</asp:ListItem>
                </asp:DropDownList>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td class="style2">
            </td>
            <td class="style2">
                <asp:Label ID="Label4" runat="server" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="New Permissions:" />
            </td>
            <td colspan="3" class="style2">
                <asp:RadioButton ID="rbAllowedToPick" Font-Names="Verdana" Font-Size="XX-Small" runat="server" GroupName="AllowedToPick" OnCheckedChanged="rbAllowedToPick_CheckedChanged" Text="Allowed to pick" AutoPostBack="True" /><asp:RadioButton ID="rbNotAllowedToPick" Font-Names="Verdana" Font-Size="XX-Small" runat="server" GroupName="AllowedToPick" Text="NOT allowed to pick" OnCheckedChanged="rbNotAllowedToPick_CheckedChanged" AutoPostBack="True" />
                &nbsp;<asp:CheckBox ID="cbApplyMaxGrab" Font-Size="XX-Small" Font-Names="Verdana" runat="server" Text="Apply MAX GRAB" Enabled="False" OnCheckedChanged="cbApplyMaxGrab_CheckedChanged" AutoPostBack="True" />
                &nbsp;
                <asp:Label ID="Label5" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Navy" Text="Quantity:" />
                <asp:TextBox ID="tbMaxGrabQty" Font-Names="Verdana" Font-Size="XX-Small" runat="server" Width="50px" Enabled="False" />
                &nbsp;<asp:RangeValidator ID="rvMaxGrabQty" runat="server" Font-Names="Verdana" Font-Size="XX-Small" ErrorMessage="value must be positive" ControlToValidate="tbMaxGrabQty" Font-Bold="True" MaximumValue="99999" MinimumValue="0" />
            </td>
            <td class="style2">
            </td>
        </tr>
        <tr>
            <td class="style1">
            </td>
            <td class="style1">
                <asp:Label ID="Label6" runat="server" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="User Group:" />
            </td>
            <td class="style1">
                <asp:DropDownList ID="ddlUserGroup" runat="server" Font-Size="X-Small" AutoPostBack="True">
                    <asp:ListItem Value="-1">- please select -</asp:ListItem>
                </asp:DropDownList>
            </td>
            <td class="style1">
            </td>
            <td class="style1">
            </td>
            <td class="style1">
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnUpdateProductPermissions" runat="server" OnClick="btnUpdateProductPermissions_Click" Text="update product permissions" Enabled="False" />
            </td>
            <td>
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
            <td>
            </td>
            <td colspan="5">
                <hr />
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td colspan="2">
                <asp:Label ID="Label7" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="Switch Customer for User Account
                    " />
            &nbsp;<asp:Label ID="Label10" runat="server" Font-Bold="True" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Red" Text="THIS FACILITY IS NOT YET AVAILABLE" />
            </td>
            <td>
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
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Label ID="Label8" runat="server" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="User account:" />
            </td>
            <td>
                <%--<asp:DropDownList ID="ddlUser" runat="server" Font-Size="X-Small" AutoPostBack="True" > <asp:ListItem Value="-1">- please select -</asp:ListItem>
                </asp:DropDownList>--%>
                <telerik:RadComboBox ID="rcbUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" OnSelectedIndexChanged="rcbUser_SelectedIndexChanged" AutoPostBack="True" HighlightTemplatedItems="true" CausesValidation="False" EnableLoadOnDemand="True" OnItemsRequested="rcbUser_ItemsRequested" EnableVirtualScrolling="True" ShowMoreResultsBox="True" Filter="Contains" ToolTip="Shows all available user IDs when no search text is specified. Search for user IDs starting with a specific character by typing that character." Enabled="False"/>
            </td>
            <td>
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
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Label ID="Label9" runat="server" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="New customer:" />
            </td>
            <td>
                <asp:DropDownList ID="ddlNewCustomer" runat="server" Font-Size="XX-Small" AutoPostBack="True" onselectedindexchanged="ddlNewCustomer_SelectedIndexChanged">
                    <asp:ListItem Value="-1">- please select -</asp:ListItem>
                </asp:DropDownList>
            &nbsp;<asp:LinkButton ID="lnkbtnNewCustomerWURS" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnNewCustomerWURS_Click">WURS</asp:LinkButton>
                &nbsp;<asp:LinkButton ID="lnkbtnNewCustomerWUIRE" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnNewCustomerWUIRE_Click">WUIRE</asp:LinkButton>
            </td>
            <td align="right">
                <asp:Label ID="lblLegendWUModelUser" runat="server" Font-Names="Verdana" Font-Size="X-Small" ForeColor="Navy" Text="Model user:" Visible="False" />
            </td>
            <td>
                <asp:DropDownList ID="ddlModelUser" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="XX-Small" OnSelectedIndexChanged="ddlModelUser_SelectedIndexChanged" Enabled="False" Visible="False" />
                <telerik:RadComboBox ID="rcbModelUser" runat="server" Font-Names="Verdana" Font-Size="XX-Small" Font-Bold="true" OnSelectedIndexChanged="rcbModelUser_SelectedIndexChanged" AutoPostBack="True" HighlightTemplatedItems="true" CausesValidation="False" EnableLoadOnDemand="True" OnItemsRequested="rcbModelUser_ItemsRequested" EnableVirtualScrolling="True" ShowMoreResultsBox="True" Filter="Contains" ToolTip="Shows what?" Enabled="False"/>
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                <asp:Button ID="btnSwitchCustomer" runat="server" Text="switch customer" Enabled="False" onclick="btnSwitchCustomer_Click" />
            </td>
            <td>
            </td>
            <td>
            </td>
            <td>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
            <td>
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
    <%--    <table>
        <tr>
            <td style="width: 442px; height: 60px">
                &nbsp;</td>
            <td style="width: 391px; height: 60px">
                &nbsp;</td>
        </tr>
        <tr>
            <td style="width: 442px; height: 60px">
                &nbsp;</td>
            <td style="width: 391px; height: 60px">
                &nbsp;</td>
        </tr>
        <tr>
            <td style="width: 442px; height: 60px">
                <br />
                <br />
                &nbsp;Customer:
                <asp:DropDownList ID="ddlCustomer" runat="server" Font-Size="X-Small" AutoPostBack="True" OnSelectedIndexChanged="ddlCustomer_SelectedIndexChanged">
                </asp:DropDownList>
                <br />
                <br />
                &nbsp;User:&nbsp;
                <asp:LinkButton ID="lnkbtnVerifySelectedUser" runat="server" onclick="lnkbtnVerifySelectedUser_Click" Enabled="False">verify selected user</asp:LinkButton>
                <br />
                <br />
                &nbsp;<br />
                UserID list (comma, space or newline separated):
                <asp:LinkButton ID="lnkbtnClearUserIDListbox" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnClearUserIDListbox_Click">clear</asp:LinkButton>
                <br />
                &nbsp;<asp:TextBox ID="tbUserIDs" runat="server" Font-Names="Arial" Font-Size="XX-Small" Width="95%" Rows="4" TextMode="MultiLine"></asp:TextBox>
                <br />
                <br />
                <asp:Button ID="btnVerify" runat="server" OnClick="btnVerify_Click" Text="Verify" Width="100px" />
                &nbsp;
                <asp:Button ID="btnCloneProfiles" runat="server" OnClick="btnCloneProfiles_Click" Text="Clone profile" Width="100px" />
                &nbsp;<asp:LinkButton ID="lnkbtnBuildSQLString" runat="server" Font-Names="Arial" Font-Size="XX-Small" OnClick="lnkbtnBuildSQLString_Click">build SQL string</asp:LinkButton>
                <br />
                <br />
                (clones from selected user to UserID list)<br />
                <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
                <br />
                <br />
            </td>
            <td style="width: 391px; height: 60px">
                &nbsp;Quit after reporting
                <asp:TextBox ID="tbQuitLevel" runat="server" Width="42px" Font-Names="Arial" Font-Size="XX-Small">10</asp:TextBox>
                affected users<br />
                <br />
                <br />
                <br />
                <asp:CheckBox ID="cbDontReportWhenEntireProfileMissing" runat="server" Text="Don't report when entire profile missing" /><br />
                <br />
                <asp:CheckBox ID="cbGenerateSQLToFixMissingEntries" runat="server" Text="Generate SQL to fix missing entries" Checked="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbGenerateSQLToMakeActive" runat="server" Text="Generate SQL to make Active" Checked="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbOverrideCloningGroupSanityCheck" runat="server" Text="Override cloning group sanity check" AutoPostBack="True" oncheckedchanged="cbOverrideCloningGroupSanityCheck_CheckedChanged" />
                <br />
                <br />
                <asp:CheckBox ID="cbIncludeSuspendedUsers" runat="server" Text="Include suspended users " AutoPostBack="True" />
                <br />
                <br />
                <asp:CheckBox ID="cbOutputDebugInformation" runat="server" Text="Output debug information" />
            </td>
        </tr>
    </table>
    Summary:<br />
    <asp:TextBox ID="tbSummary" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    Detail:<br />
    <asp:TextBox ID="tbDetail" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    SQL:<br />
    <asp:TextBox ID="tbSQL" runat="server" Height="224px" TextMode="MultiLine" Width="100%" Font-Names="Arial" Font-Size="XX-Small" /><br />
    <asp:LinkButton ID="lnkbtnHideProductProfile" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnHideProductProfile_Click">hide product profile</asp:LinkButton>
    <br />
    <asp:GridView ID="gvProductProfile" runat="server" CellPadding="2" Font-Names="Arial" Font-Size="XX-Small" Width="100%">
    </asp:GridView>
    <br />
    &nbsp;<asp:Button ID="btnClear" runat="server" OnClick="btnClear_Click" Text="Clear" />
    &nbsp;<asp:LinkButton ID="lnkbtnToggleHelp" runat="server" Font-Names="Arial" Font-Size="XX-Small" onclick="lnkbtnToggleHelp_Click">toggle help</asp:LinkButton>
    --%>
    <asp:Panel ID="pnlHelp" runat="server" Width="100%" Font-Names="Verdana" Font-Size="X-Small" Visible="false">
        <br />
        <strong>HELP</strong>
        <br />
        <br />
        </asp:Panel>
    </form>
</body>
</html>
